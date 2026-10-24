# WhatsApp Messages Sender — v0.6

A GUI + CLI tool for sending WhatsApp messages (with or without an image) to contacts stored in an Excel file (`data.xlsx`), powered by `pywhatkit`.

- **CLI (`main.py`)** — send messages, send messages with an image, add/edit contacts (launches the GUI), about screen.
- **GUI (`gui.py`)** — add, edit, delete, preview and restore contacts with full Arabic (RTL) text support.
- **Config (`config.ini`)** — runtime settings for sending and the GUI.

---

## Requirements

- Python 3.x
- `customtkinter`, `openpyxl`, `pywhatkit`, `arabic_reshaper`, `python-bidi`, `rich`

---

## Config (`config.ini`)

```ini
[settings]
load_data = data.xlsx   # Contacts workbook (default = data.xlsx)
wait_time = 10          # Seconds to wait before sending (default = 10)
tab_close = True        # Close browser tab after sending (default = True)
image_folder = images/  # Folder containing the image to send
page_size = 25          # Rows per page in the GUI preview (default = 25)
```

---

# Changes: v0.4 → v0.5

## `main.py`

| Area | v0.4 | v0.5 |
|---|---|---|
| Version | `0.4 beta` (20 Oct, 2026) | `0.5 beta` (21 Oct, 2026) |
| Config | Module-level `configparser` reads only `wait_time`, `tab_close`, `image_folder` | `config_file()` function returning a dict (`load_data`, `wait_time`, `tab_close`, `image_folder`, `page_size`), exported as module-level globals; supports inline `#`/`;` comments |
| Contacts file | Hardcoded `data.xlsx` | Configurable via `load_data` setting |
| Console output | Plain `print` | `rich.print` with color markup (`[blue]`, `[red]`, `[green]`, `[cyan]`) |
| Menu | `[1] Send`, `[2] Send with image`, `[3] About`, `[0] Exit` | Adds `[3] Add / Edit Contacts` (launches `gui.py`), About moved to `[4]` |
| Input | Plain `input()` for customer count | New `input_with_esc()` — Esc cancels and returns to the menu, Backspace edits, works on Windows and Linux |
| Robustness | Customer count not validated | `call_customer_info_func()` handles Esc/cancel before proceeding |

## `gui.py`

### Arabic text editing (`ArabicEntry`)
- **v0.4:** Append-only editing — Backspace deleted the last raw character regardless of cursor position; mid-text typing, selection, Delete and undo/redo did not work correctly with RTL text.
- **v0.5:** Full bidi display↔raw index mapping:
  - `_display_from_raw`, `_disp_to_raw`, `_raw_to_disp`, `_backspace_raw_index`, `_delete_raw_index` — correct cursor/character mapping for RTL text.
  - Mid-text typing inserts at the correct raw position.
  - Backspace / Delete delete the correct character (before/after cursor).
  - Selection support: select a range and type/Backspace/Delete replaces it.
  - **Ctrl+C / Ctrl+X / Ctrl+V** copy/cut/paste raw text (paste inserts at cursor, replaces selection).
  - **Ctrl+A** select-all, **Ctrl+Z/Y** blocked (undo would corrupt raw/display state).
  - Ligatures disabled in the entry (`_entry_reshaper`) so cursor mapping stays 1:1; still enabled in labels (`ar()`).
  - New `set(text)` method for pre-filling (used by the edit window).
  - Graceful fallback when `arabic_reshaper`/`python-bidi` are missing.

### Add window
- Name/phone entries centered (`justify="center"`).
- Buttons use a consistent dark gray (`#444444`) style.

### Edit system (new)
- **`EditWindow`** — dialog to edit an existing contact's name and phone, saving back to the exact Excel row (`ws.cell(row=...)`).
- **Edit button** on every row in the data view.

### Delete (new)
- **Delete button** on every row in the data view (`delete_row()`); verifies the row is non-empty before removing it.

### Finished-sheet view (new)
- `FINISHED_SHEET = 'finished'` is created alongside `data` if missing.
- **Data / Finished segmented toggle** — switch between the pending and finished sheets.
- **Restore button** in the Finished view — moves a contact back to the `data` sheet.

### Preview window
- Window `580x480`, `minsize(600, 500)`; resizable state set explicitly.
- **Line number (`#`)** column showing global row numbers.
- **Refresh button moved to the top-right** of the title row (was below the list).
- **Actions column header** appears after data loads.

### Performance
- **Pagination** — only `page_size` rows (config, default 25) rendered at a time; Prev/Next page buttons with disabled states at the ends.
- **Page number entry** — type a page number and press **Enter** to jump (`_goto_page`); the separate "Go" button was removed.
- **Background loading** — `threading.Thread` + `openpyxl.load_workbook(read_only=True)` so the UI stays responsive; contacts cached in `self._contacts` and pages rendered from the cache.
- **Loading animation** — indeterminate `CTkProgressBar` + "Loading contacts..." label shown while loading; list/pagination hidden during load.
- Workbook loading guards against missing sheets/failed reads; page clamped after refresh.

### Config
- `gui.py` reads `config.ini` directly via its own `_load_config()` (no import from `main`), providing `EXCEL_FILE` (`load_data`) and `PAGE_SIZE` (`page_size`).

## `config.ini`
- Added `load_data` (contacts workbook path) and `page_size` (GUI rows per page).
- All settings now documented with inline comments; removed stray quotes around `image_folder`.

## Files
| | v0.4 | v0.5 |
|---|---|---|
| `main.py` | 193 lines | 277 lines |
| `gui.py` | 210 lines | 681 lines |
| `config.ini` | 3 settings | 5 settings |

# Changes: v0.5 → v0.6

## `main.py`

| Area | v0.5 | v0.6 |
|---|---|---|
| Version | `0.5 beta` (21 Oct, 2026) | `0.6 beta` (24 Oct, 2026) |
| Phone validation | None | Added `_normalize_phone()` and `_is_valid_phone()` helpers; warnings shown for invalid phone numbers when sending |

## `gui.py`

### Phone validation
- **Normalization:** Leading `0` stripped during normalization; phone numbers stored in normalized form.
- **Validation:** Rejects invalid phone numbers (must be valid digit lengths). Invalid entries show "Phone number must be 10 digits" warning.
- **Duplicate prevention:** Checks for existing normalized phone numbers on add/edit and prevents duplicates.
- **Safe I/O:** Workbooks closed properly after operations.

### Edit window redesign
- **v0.5:** Simple framed layout.
- **v0.6:** Modern card-style design with header bar, stacked labels/inputs, improved button styling and layout (`Save`/`Cancel`).

### Preview window
- **Delete Duplicates:** New button in the top-left corner (Data view only), appears when duplicates exist. Deletes duplicate numbers (keeps first occurrence).
- **Icons:** Added icons/symbols to action buttons (Add, Preview Data, Save, Cancel, Refresh, Edit, Delete, Restore, duplicates).
- **View toggle:** "Finished" renamed to "Sent".
- **Navigation:** Page number entry supports `Enter` and `Keypad Enter`.
- **UI tweaks:** Improved button styling/spacing, title alignment adjusted to accommodate new button.

### General
- Better phone number handling across add/edit/duplicate logic.
