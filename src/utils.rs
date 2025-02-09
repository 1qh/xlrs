use crate::types::{App, ColumnState};
use comfy_table::{presets::UTF8_FULL, ContentArrangement, Table};
use crossterm::event::KeyCode::{self, Down, Up};
use ratatui::{
  layout::{Constraint, Direction, Layout, Rect},
  Frame,
};
use unicode_width::UnicodeWidthStr;
use unidecode::unidecode;

pub fn normalize_text(text: &str) -> String {
  unidecode(text).replace([' ', '-'], "_").to_lowercase()
}

pub fn navigate_index(cur: usize, len: usize, key: KeyCode) -> usize {
  match key {
    Down => (cur + 1) % len,
    Up => (cur + len - 1) % len,
    _ => cur,
  }
}

pub fn visual_width(text: &str) -> usize {
  UnicodeWidthStr::width(text)
}

pub fn layout(f: &mut Frame) -> (Rect, Rect, Rect) {
  let chunks = Layout::default()
    .direction(Direction::Vertical)
    .constraints([Constraint::Length(1), Constraint::Min(1), Constraint::Length(1)])
    .split(f.area());
  (chunks[0], chunks[1], chunks[2])
}

pub fn create_table(app: &App, width: u16) -> Table {
  let mut table = Table::new();
  table
    .load_preset(UTF8_FULL)
    .set_content_arrangement(ContentArrangement::Dynamic)
    .set_width(width);

  let visible_data = app
    .data
    .iter()
    .skip(app.first_row)
    .filter(|row| app.is_row_visible(row))
    .skip(app.current_page * app.rows_per_page)
    .take(app.rows_per_page);

  for row in visible_data {
    let visible_cells = row
      .iter()
      .enumerate()
      .filter(|(i, _)| matches!(app.columns[*i], ColumnState::Original | ColumnState::NonEmpty))
      .map(|(_, cell)| cell.to_string())
      .collect::<Vec<_>>();
    table.add_row(visible_cells);
  }
  table
}
