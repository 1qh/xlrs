mod types;
mod utils;
use calamine::{open_workbook, Reader, Xlsx};
use crossterm::{
  event::{
    self, poll,
    Event::Key,
    KeyCode::{Char, Down, End, Enter, Home, Left, Right, Tab, Up},
    KeyEvent, KeyModifiers,
  },
  execute,
  terminal::{disable_raw_mode, enable_raw_mode, EnterAlternateScreen, LeaveAlternateScreen},
};
use ratatui::{
  backend::CrosstermBackend,
  layout::{Alignment, Constraint, Direction::Vertical, Layout},
  style::{Color, Style},
  text::{Line, Span},
  widgets::{List, ListItem, Paragraph},
  Frame, Terminal,
};
use serde_json::to_string_pretty;
use std::{
  collections::{HashMap, HashSet},
  env::args,
  error::Error,
  fs::File,
  io::{stdout, Read, Seek, Write},
  time::{Duration, Instant},
};
use tui_input::{backend::crossterm::EventHandler, Input};

use types::{App, ColumnConfig, ColumnState, ExportEdit, Step, FOCUSED_STYLE};
use utils::{create_table, layout, navigate_index, normalize_text, visual_width};

impl App {
  fn load_sheet<T: Read + Seek>(&mut self, xlsx: &mut Xlsx<T>) -> bool {
    if let Some(idx) = self.selected_sheet {
      if let Ok(range) = xlsx.worksheet_range(&self.sheets[idx]) {
        self.data =
          range.rows().map(|row| row.iter().map(|cell| cell.to_string()).collect()).collect();
        if self.data.is_empty() {
          return false;
        }
        let col_count = self.data.get(0).map_or(0, |r| r.len());
        self.columns = vec![ColumnState::Hidden; col_count];
        self.column_configs = vec![ColumnConfig::default(); col_count];
        self.custom_keys = vec![Input::default(); col_count];
        self.export_filename = {
          let default = self.get_default_filename();
          Input::default().with_value(default)
        };
        return true;
      }
    }
    false
  }
  fn is_row_visible(&self, row: &[String]) -> bool {
    row
      .iter()
      .enumerate()
      .filter(|(i, _)| matches!(self.columns[*i], ColumnState::Original | ColumnState::NonEmpty))
      .all(|(i, cell)| matches!(self.columns[i], ColumnState::Original) || !cell.trim().is_empty())
  }
  fn handle_key(
    &mut self,
    key: KeyEvent,
    modifiers: KeyModifiers,
    xlsx: &mut Xlsx<impl Read + Seek>,
  ) -> bool {
    if modifiers.contains(KeyModifiers::CONTROL) {
      match key.code {
        Char('q') => return true,
        Char('b') => {
          self.step = self.handle_back();
        }
        _ => {}
      }
    }
    match self.step {
      Step::SheetSelect => self.handle_sheet_select(key, xlsx),
      Step::RowTrim => self.handle_row_trim(key, xlsx),
      Step::MergePrompt => self.handle_merge_prompt(key, xlsx),
      Step::ColSelect => self.handle_col_select(key),
      Step::Preview => self.handle_preview(key),
      Step::Export => self.handle_export(key),
    }
    false
  }
  fn handle_sheet_select<T: Read + Seek>(&mut self, key: KeyEvent, xlsx: &mut Xlsx<T>) {
    match key.code {
      Up => {
        self.selected_sheet = self.selected_sheet.map(|i| i.saturating_sub(1)).or(Some(0));
      }
      Down => {
        self.selected_sheet =
          self.selected_sheet.map(|i| (i + 1).min(self.sheets.len() - 1)).or(Some(0));
      }
      Enter => {
        if self.load_sheet(xlsx) {
          self.step = Step::RowTrim;
        }
      }
      _ => {
        self.sheet_search.handle_event(&Key(key));
        self.update_sheet_search();
      }
    }
  }
  fn handle_row_trim<T: Read + Seek>(&mut self, key: KeyEvent, xlsx: &mut Xlsx<T>) {
    match key.code {
      Enter => {
        if let Ok(row) = self.row_input.value().trim().parse::<usize>() {
          if row < self.data.len() {
            self.first_row = row;
            let merge_info = self.check_merge_options(xlsx);
            if !merge_info.is_empty() {
              self.merge_info = Some(merge_info);
              self.step = Step::MergePrompt;
            } else {
              self.step = Step::ColSelect;
            }
          }
        }
      }
      _ => {
        self.row_input.handle_event(&Key(key));
      }
    }
  }
  fn handle_merge_prompt<T: Read + Seek>(&mut self, key: KeyEvent, xlsx: &mut Xlsx<T>) {
    match key.code {
      Char('y') | Char('Y') => {
        self.perform_merge(xlsx);
        self.merge_info = None;
        self.step = Step::ColSelect;
      }
      Char('n') | Char('N') => {
        self.merge_info = None;
        self.step = Step::ColSelect;
      }
      _ => {}
    }
  }
  fn handle_col_select(&mut self, key: KeyEvent) {
    match key.code {
      Char(' ') => self.toggle_col_select(),
      Char('a') => self.toggle_all_col(),
      Up | Down => {
        self.selected_column = navigate_index(self.selected_column, self.columns.len(), key.code);
      }
      Enter => {
        self.selected_column = 0;
        self.step = Step::Preview;
      }
      _ => {}
    }
  }
  fn handle_preview(&mut self, key: KeyEvent) {
    match key.code {
      Left => self.prev_page(),
      Right => self.next_page(),
      Char(' ') => self.toggle_col_filter(),
      Char('+') => {
        self.rows_per_page = self.rows_per_page.saturating_add(1);
        self.current_page = 0;
      }
      Char('-') => {
        self.rows_per_page = self.rows_per_page.saturating_sub(1).max(5);
        self.current_page = 0;
      }
      Home => {
        self.current_page = 0;
      }
      End => {
        self.current_page = self.total_pages().saturating_sub(1);
      }
      Up | Down => {
        let visible = self.visible_columns();
        if !visible.is_empty() {
          self.selected_column = navigate_index(self.selected_column, visible.len(), key.code);
        }
      }
      Enter => {
        self.step = Step::Export;
        self.export_focus_row = 0;
        self.export_edit = ExportEdit::FileName;
      }
      _ => {}
    }
  }
  fn get_export_target(&mut self) -> Option<&mut Input> {
    let visible = self.visible_columns();
    let col_idx = visible.get(self.export_focus_row.checked_sub(1)?)?;
    match self.export_edit {
      ExportEdit::KeyStr => Some(&mut self.custom_keys[*col_idx]),
      ExportEdit::Prefix => Some(&mut self.column_configs[*col_idx].prefix),
      ExportEdit::Postfix => Some(&mut self.column_configs[*col_idx].postfix),
      ExportEdit::FileName | ExportEdit::Deduplicate => None,
    }
  }
  fn handle_export(&mut self, key: KeyEvent) {
    match key.code {
      Up | Down => {
        let total_items = self.visible_columns().len() + 2;
        let delta = if key.code == Down { 1 } else { total_items - 1 };
        self.export_focus_row = (self.export_focus_row + delta) % total_items;
        self.export_edit =
          if self.export_focus_row == 0 { ExportEdit::FileName } else { ExportEdit::KeyStr };
      }
      Tab => {
        self.export_edit = if self.export_focus_row == 0 {
          match self.export_edit {
            ExportEdit::FileName => ExportEdit::Deduplicate,
            _ => ExportEdit::FileName,
          }
        } else {
          match self.export_edit {
            ExportEdit::KeyStr => ExportEdit::Prefix,
            ExportEdit::Prefix => ExportEdit::Postfix,
            _ => ExportEdit::KeyStr,
          }
        };
      }
      Enter => self.export_to_json(),
      _ => {
        match self.export_focus_row {
          0 => match self.export_edit {
            ExportEdit::FileName => {
              self.export_filename.handle_event(&Key(key));
            }
            ExportEdit::Deduplicate if key.code == Char(' ') => {
              self.deduplicate ^= true;
            }
            _ => {}
          },
          _ => {
            if let Some(target) = self.get_export_target() {
              target.handle_event(&Key(key));
            }
          }
        };
      }
    }
  }
  fn next_page(&mut self) {
    if self.current_page + 1 < self.total_pages() {
      self.current_page += 1;
    }
  }
  fn prev_page(&mut self) {
    if self.current_page > 0 {
      self.current_page -= 1;
    }
  }
  fn toggle_col_filter(&mut self) {
    if let Some(&col_idx) = self.visible_columns().get(self.selected_column) {
      self.columns[col_idx] = match self.columns[col_idx] {
        ColumnState::Original => ColumnState::NonEmpty,
        _ => ColumnState::Original,
      };
    }
  }
  fn toggle_col_select(&mut self) {
    if let Some(col) = self.columns.get_mut(self.selected_column) {
      *col = match *col {
        ColumnState::Hidden => ColumnState::NonEmpty,
        _ => ColumnState::Hidden,
      };
    }
  }
  fn toggle_all_col(&mut self) {
    let all_hidden = self.columns.iter().all(|&c| matches!(c, ColumnState::Hidden));
    self.columns.fill(if all_hidden { ColumnState::NonEmpty } else { ColumnState::Hidden });
  }
  fn update_sheet_search(&mut self) {
    let search_lower = self.sheet_search.value().to_lowercase();
    self.matching_sheets = self
      .sheets
      .iter()
      .enumerate()
      .filter_map(|(i, name)| (name.to_lowercase().contains(&search_lower)).then(|| i))
      .collect();
    self.selected_sheet = self.matching_sheets.first().copied().or(self.selected_sheet);
  }
  fn input_style(&self, is_selected: bool, edit_mode: ExportEdit) -> Style {
    if !is_selected {
      return Style::default();
    }
    if self.export_edit == edit_mode {
      FOCUSED_STYLE.bg(Color::DarkGray)
    } else {
      FOCUSED_STYLE
    }
  }
  fn export_to_json(&mut self) {
    let filename = if !self.export_filename.value().is_empty() {
      normalize_text(self.export_filename.value())
    } else {
      self.get_default_filename()
    };
    let filepath = format!("{}.json", filename);
    if let Ok(mut file) = File::create(&filepath) {
      let records = self.create_json_records();
      let records = if self.deduplicate {
        let mut seen = HashSet::new();
        records.into_iter().filter(|rec| seen.insert(serde_json::to_string(rec).unwrap())).collect()
      } else {
        records
      };
      let json_array = serde_json::Value::Array(records);
      if writeln!(file, "{}", to_string_pretty(&json_array).unwrap()).is_ok() {
        self.export_toast = Some(format!("Exported to {}.json successfully", filename));
        self.export_toast_time = Some(Instant::now());
      }
    }
  }
  fn check_merge_options<T: Read + Seek>(&self, xlsx: &mut Xlsx<T>) -> Vec<(String, Vec<String>)> {
    let mut info = Vec::new();
    let primary_header = match self.data.get(self.first_row) {
      Some(row) => row,
      None => return info,
    };
    for (i, sheet_name) in self.sheets.iter().enumerate() {
      if Some(i) == self.selected_sheet {
        continue;
      }
      if let Ok(range) = xlsx.worksheet_range(sheet_name) {
        if let Some(header_row) = range.rows().nth(self.first_row) {
          let sheet_set: HashSet<_> =
            header_row.iter().map(|s| s.to_string().trim().to_string()).collect();
          let mutual: Vec<String> = primary_header
            .iter()
            .map(|s| s.trim().to_string())
            .filter(|s| sheet_set.contains(s))
            .collect();
          if !mutual.is_empty() {
            info.push((sheet_name.clone(), mutual));
          }
        }
      }
    }
    info
  }
  fn perform_merge<T: Read + Seek>(&mut self, xlsx: &mut Xlsx<T>) {
    let primary_header = match self.data.get(self.first_row) {
      Some(row) => row,
      None => return,
    };

    let mut common: HashSet<String> = primary_header.iter().map(|s| s.trim().to_string()).collect();
    if let Some(ref info) = self.merge_info {
      for (_, mutual) in info {
        let sheet_set: HashSet<String> = mutual.iter().cloned().collect();
        common = common.into_iter().filter(|s| sheet_set.contains(s)).collect();
      }
    }
    let new_header: Vec<String> =
      primary_header.iter().filter(|s| common.contains(&s.trim().to_string())).cloned().collect();
    let mut merged_data = Vec::new();
    merged_data.push(new_header.clone());
    let mut merge_sheet = |sheet_name: &String| {
      if let Ok(range) = xlsx.worksheet_range(sheet_name) {
        let rows: Vec<_> = range.rows().collect();
        if rows.len() <= self.first_row {
          return;
        }
        let header_row = rows[self.first_row];
        let header_map: HashMap<String, usize> = header_row
          .iter()
          .enumerate()
          .map(|(idx, cell)| (cell.to_string().trim().to_string(), idx))
          .collect();
        for row in rows.iter().skip(self.first_row + 1) {
          let new_row: Vec<String> = new_header
            .iter()
            .map(|col_name| {
              if let Some(&idx) = header_map.get(&col_name.trim().to_string()) {
                row.get(idx).map(|s| s.to_string()).unwrap_or_default()
              } else {
                String::new()
              }
            })
            .collect();
          merged_data.push(new_row);
        }
      }
    };
    merge_sheet(&self.sheets[self.selected_sheet.unwrap()]);
    if let Some(ref info) = self.merge_info {
      for (sheet_name, _) in info {
        merge_sheet(sheet_name);
      }
    }
    self.data = merged_data;
    self.sheets = vec!["[Merged]".to_string()];
    self.selected_sheet = Some(0);
    self.first_row = 0;
    self.export_filename = Input::default().with_value(self.get_default_filename());
  }
}
fn ui(f: &mut Frame, app: &mut App) {
  if let Some(time) = app.export_toast_time {
    if time.elapsed().as_secs() >= 3 {
      app.export_toast = None;
      app.export_toast_time = None;
    }
  }
  let (header, content, footer) = layout(f);

  match app.step {
    Step::MergePrompt => {
      let mut lines = vec![];
      if let Some(ref info) = app.merge_info {
        lines
          .push(format!("Merge data from other sheets?\n\n      Sheet        Mutual columns\n",));
        for (sheet, mutual) in info {
          lines.push(format!("{:<16} | {}", sheet, mutual.join(", ")));
        }
      }
      let para = Paragraph::new(lines.join("\n"));
      f.render_widget(para, f.area());
    }
    Step::SheetSelect => {
      let label = "Search: ";
      let input_text = format!("{}{}", label, app.sheet_search.value());
      f.set_cursor_position((
        header.x + (label.len() + app.sheet_search.value().len()) as u16,
        header.y,
      ));
      f.render_widget(Paragraph::new(input_text), header);

      let items: Vec<ListItem> = if app.sheet_search.value().is_empty() {
        app
          .sheets
          .iter()
          .enumerate()
          .map(|(i, sheet)| {
            ListItem::new(sheet.as_str()).style(if Some(i) == app.selected_sheet {
              FOCUSED_STYLE
            } else {
              Style::default()
            })
          })
          .collect()
      } else {
        app
          .matching_sheets
          .iter()
          .map(|&i| {
            ListItem::new(app.sheets[i].as_str()).style(if Some(i) == app.selected_sheet {
              FOCUSED_STYLE
            } else {
              Style::default()
            })
          })
          .collect()
      };
      let list = List::new(items);
      f.render_widget(list, content);
    }
    Step::RowTrim => {
      let label = "First row number: ";
      let input_text = format!("{}{}", label, app.row_input.value());
      f.set_cursor_position((
        header.x + (label.len() + app.row_input.value().len()) as u16,
        header.y,
      ));
      f.render_widget(Paragraph::new(input_text), header);

      let preview: Vec<String> = app
        .data
        .iter()
        .enumerate()
        .map(|(i, row)| format!("{:<2} | {}", i, row.join(", ")))
        .collect();
      let para = Paragraph::new(preview.join("\n"));
      f.render_widget(para, content);
    }
    Step::ColSelect => {
      let columns: Vec<Line> = app
        .data
        .get(app.first_row)
        .map(|row| {
          row
            .iter()
            .enumerate()
            .map(|(i, col)| {
              let style = if i == app.selected_column { FOCUSED_STYLE } else { Style::default() };
              Line::styled(
                format!(
                  "  {} {}",
                  match app.columns[i] {
                    ColumnState::Hidden => "◯",
                    _ => "●",
                  },
                  col
                ),
                style,
              )
            })
            .collect()
        })
        .unwrap_or_default();
      let text = columns.into_iter().collect::<Vec<Line>>();
      f.set_cursor_position((0, f.area().y + 1 + app.selected_column as u16));
      let para = Paragraph::new(text);
      f.render_widget(para, content);
    }
    Step::Preview => {
      let chunks = Layout::default()
        .direction(Vertical)
        .constraints([Constraint::Length(app.visible_columns().len() as u16), Constraint::Min(0)])
        .split(f.area());

      let filter_info = app
        .visible_columns()
        .iter()
        .enumerate()
        .map(|(i, &col_idx)| {
          let style = if i == app.selected_column { FOCUSED_STYLE } else { Style::default() };
          let column_name = app
            .data
            .get(app.first_row)
            .and_then(|row| row.get(col_idx))
            .map(|s| s.as_str())
            .unwrap_or("Unknown");
          Line::styled(
            format!(
              "  {} · {}",
              match app.columns[col_idx] {
                ColumnState::Original => "◯ Original",
                ColumnState::NonEmpty => "● NonEmpty",
                ColumnState::Hidden => "Hidden",
              },
              column_name
            ),
            style,
          )
        })
        .collect::<Vec<_>>();
      f.set_cursor_position((0, app.selected_column as u16));
      f.render_widget(Paragraph::new(filter_info), chunks[0]);
      f.render_widget(Paragraph::new(create_table(app, f.area().width).to_string()), chunks[1]);
    }
    Step::Export => {
      let visible_columns = app.visible_columns();
      let name_col_width = visible_columns
        .iter()
        .map(|&col_idx| visual_width(&app.data[app.first_row][col_idx]))
        .max()
        .unwrap_or(20)
        .max(20)
        + 1;

      let filename_style = app.input_style(
        app.export_focus_row == 0 && app.export_edit == ExportEdit::FileName,
        ExportEdit::FileName,
      );
      let dedup_style = app.input_style(
        app.export_focus_row == 0 && app.export_edit == ExportEdit::Deduplicate,
        ExportEdit::Deduplicate,
      );
      let dedup_box = if app.deduplicate { " ● " } else { " ◯ " };

      let line0 = Line::from(vec![
        Span::raw("Filename: "),
        Span::styled(format!("{}", app.export_filename), filename_style),
        Span::raw("   Deduplicate "),
        Span::styled(dedup_box, dedup_style),
      ]);
      f.render_widget(Paragraph::new(line0), header);

      let mut lines = vec![Line::raw("")];
      for (i, &col_idx) in visible_columns.iter().enumerate() {
        let config = &app.column_configs[col_idx];
        let column_name = &app.data[app.first_row][col_idx];
        let is_selected = app.export_focus_row > 0 && i == (app.export_focus_row - 1);
        let custom_key = if app.custom_keys[col_idx].value().is_empty() {
          normalize_text(column_name)
        } else {
          normalize_text(app.custom_keys[col_idx].value())
        };
        let display_name = match app.export_edit {
          ExportEdit::KeyStr if is_selected => format!("{}", custom_key),
          _ => custom_key,
        };
        let display_prefix = match app.export_edit {
          ExportEdit::Prefix if is_selected => format!("{}", config.prefix.value()),
          _ => config.prefix.value().to_string(),
        };
        let display_postfix = match app.export_edit {
          ExportEdit::Postfix if is_selected => format!("{}", config.postfix.value()),
          _ => config.postfix.value().to_string(),
        };
        let spans = vec![
          Span::styled(
            format!("{:<name_col_width$}", column_name),
            app.input_style(is_selected, ExportEdit::FileName),
          ),
          Span::styled(" key: ", app.input_style(is_selected, ExportEdit::FileName)),
          Span::styled(
            format!("{:<name_col_width$}", display_name),
            app.input_style(is_selected, ExportEdit::KeyStr),
          ),
          Span::styled(" prefix: ", app.input_style(is_selected, ExportEdit::FileName)),
          Span::styled(
            format!("{:<20}", display_prefix),
            app.input_style(is_selected, ExportEdit::Prefix),
          ),
          Span::styled(" postfix: ", app.input_style(is_selected, ExportEdit::FileName)),
          Span::styled(
            format!("{:<20}", display_postfix),
            app.input_style(is_selected, ExportEdit::Postfix),
          ),
        ];
        lines.push(Line::from(spans));
      }

      f.render_widget(Paragraph::new(lines), content);
      if let Some(msg) = &app.export_toast {
        f.render_widget(Paragraph::new(msg.as_str()).alignment(Alignment::Center), footer);
      }
      f.set_cursor_position((0, 100));
    }
  }

  let navigate_guide = "↑↓ to navigate";
  let toggle_guide = "Space to toggle";
  let back_guide = "Ctrl+B to go back";
  let quit_guide = "Ctrl+Q to quit";
  let export_guide = "Enter to export";

  let footer_text = if let Some(msg) = &app.export_toast {
    msg.to_string()
  } else {
    match app.step {
      Step::SheetSelect => format!("{} · {}", navigate_guide, quit_guide),
      Step::RowTrim => format!("{} · {}", back_guide, quit_guide),
      Step::ColSelect => {
        format!("{} · {} · 'a' to toggle all · {}", navigate_guide, toggle_guide, quit_guide)
      }
      Step::Preview => format!(
        "{} · {} · Page ←{}/{}→ · Rows/page: -{}+ · {}",
        navigate_guide,
        toggle_guide,
        app.current_page + 1,
        app.total_pages().max(1),
        app.rows_per_page,
        export_guide
      ),
      Step::Export => format!("{} · Tab to cycle fields · {}", navigate_guide, export_guide),
      Step::MergePrompt => "y/n".to_string(),
    }
  };
  f.render_widget(Paragraph::new(footer_text).alignment(Alignment::Center), footer);
}

fn main() -> Result<(), Box<dyn Error>> {
  let args: Vec<String> = args().collect();
  if args.len() != 2 {
    println!("Usage: {} <excel_file>", args[0]);
    return Ok(());
  }
  let mut xlsx = open_workbook(&args[1])?;
  let mut app = App::new(&mut xlsx, &args[1]);
  enable_raw_mode()?;
  let mut stdout = stdout();
  execute!(stdout, EnterAlternateScreen)?;
  let backend = CrosstermBackend::new(stdout);
  let mut terminal = Terminal::new(backend)?;

  loop {
    if poll(Duration::from_millis(100))? {
      if let Key(key) = event::read()? {
        if app.handle_key(key, key.modifiers, &mut xlsx) {
          break;
        }
      }
    }
    terminal.draw(|f| ui(f, &mut app))?;
    terminal.show_cursor()?;
  }
  disable_raw_mode()?;
  execute!(terminal.backend_mut(), LeaveAlternateScreen)?;
  Ok(())
}
