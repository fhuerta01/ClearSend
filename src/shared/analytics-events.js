// Closed vocabulary shared by client and endpoint. Never add Outlook-derived properties.
const EVENTS = Object.freeze([
  "pane_open",
  "process_click",
  "process_success",
  "process_blocked",
  "process_error",
  "quick_clean_click",
  "quick_clean_success",
  "quick_clean_blocked",
  "quick_clean_error",
  "undo_click",
  "export_click",
  "export_invalid_click",
  "remove_click",
  "copy_click",
  "refresh_click",
  "settings_click",
]);
module.exports = { EVENTS };
