use calamine::{Reader, Xlsx};
use ratatui::style::{Color, Modifier, Style};
use std::{
  io::{Read, Seek},
  path::Path,
  time::Instant,
};
use tui_input::Input;

use crate::utils::normalize_text;
pub const FOCUSED_STYLE: Style = Style::new().fg(Color::Yellow).add_modifier(Modifier::BOLD);

#[derive(Copy, Clone)]
pub enum ColumnState {
  Hidden,
  Original,
  NonEmpty,
}

#[derive(Clone)]
pub struct ColumnConfig {
  pub prefix: Input,
  pub postfix: Input,
}

impl Default for ColumnConfig {
  fn default() -> Self {
    Self { prefix: Input::default(), postfix: Input::default() }
  }
}

pub struct App {
  pub sheets: Vec<String>,
  pub selected_sheet: Option<usize>,
  pub data: Vec<Vec<String>>,
  pub first_row: usize,
  pub columns: Vec<ColumnState>,
  pub selected_column: usize,
  pub step: Step,
  pub row_input: Input,
  pub current_page: usize,
  pub rows_per_page: usize,
  pub sheet_search: Input,
  pub matching_sheets: Vec<usize>,
  pub column_configs: Vec<ColumnConfig>,
  pub export_focus_row: usize,
  pub export_edit: ExportEdit,
  pub custom_keys: Vec<Input>,
  pub export_filename: Input,
  pub export_toast: Option<String>,
  pub export_toast_time: Option<Instant>,
  pub merge_info: Option<Vec<(String, Vec<String>)>>,
  pub deduplicate: bool,
  pub original_filename: String,
}

impl App {
  pub fn new<T: Read + Seek>(xlsx: &mut Xlsx<T>, original_filename: &str) -> Self {
    let sheets = xlsx.sheet_names().to_owned();
    Self {
      sheets,
      selected_sheet: Some(0),
      data: Vec::new(),
      first_row: 0,
      columns: Vec::new(),
      selected_column: 0,
      step: Step::SheetSelect,
      row_input: Input::default(),
      current_page: 0,
      rows_per_page: 10,
      sheet_search: Input::default(),
      matching_sheets: Vec::new(),
      column_configs: Vec::new(),
      export_focus_row: 0,
      export_edit: ExportEdit::FileName,
      custom_keys: Vec::new(),
      export_filename: Input::default().with_value("export".to_string()),
      export_toast: None,
      export_toast_time: None,
      merge_info: None,
      deduplicate: true,
      original_filename: original_filename.to_string(),
    }
  }
  pub fn handle_back(&self) -> Step {
    match self.step {
      Step::SheetSelect => Step::SheetSelect,
      Step::Export => Step::Preview,
      Step::Preview => Step::ColSelect,
      Step::ColSelect | Step::MergePrompt => Step::RowTrim,
      Step::RowTrim => Step::SheetSelect,
    }
  }
  pub fn visible_columns(&self) -> Vec<usize> {
    self
      .columns
      .iter()
      .enumerate()
      .filter(|(_, &state)| matches!(state, ColumnState::Original | ColumnState::NonEmpty))
      .map(|(i, _)| i)
      .collect()
  }
  pub fn total_pages(&self) -> usize {
    let visible_rows =
      self.data.iter().skip(self.first_row).filter(|row| self.is_row_visible(row)).count();
    (visible_rows + self.rows_per_page - 1) / self.rows_per_page
  }
  pub fn get_default_filename(&self) -> String {
    if self.sheets.first().map(|s| s.as_str()) == Some("[Merged]") {
      Path::new(&self.original_filename)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or(&self.original_filename)
        .to_string()
    } else {
      self
        .selected_sheet
        .map(|idx| normalize_text(&self.sheets[idx]))
        .unwrap_or_else(|| "export".to_string())
    }
  }
  pub fn create_json_records(&self) -> Vec<serde_json::Value> {
    let visible_columns = self.visible_columns();
    self
      .data
      .iter()
      .skip(self.first_row + 1)
      .filter(|row| self.is_row_visible(row))
      .map(|row| {
        visible_columns
          .iter()
          .filter_map(|&col_idx| {
            let value = row.get(col_idx)?;
            let config = &self.column_configs[col_idx];
            let field_name = if self.custom_keys[col_idx].value().is_empty() {
              &self.data[self.first_row][col_idx]
            } else {
              self.custom_keys[col_idx].value()
            };
            Some((
              normalize_text(field_name),
              serde_json::Value::String(format!(
                "{}{}{}",
                config.prefix.value(),
                value,
                config.postfix.value()
              )),
            ))
          })
          .collect::<serde_json::Map<String, serde_json::Value>>()
          .into()
      })
      .collect()
  }
}

#[derive(PartialEq, Copy, Clone)]
pub enum Step {
  SheetSelect,
  RowTrim,
  MergePrompt,
  ColSelect,
  Preview,
  Export,
}

#[derive(PartialEq)]
pub enum ExportEdit {
  FileName,
  Deduplicate,
  KeyStr,
  Prefix,
  Postfix,
}
