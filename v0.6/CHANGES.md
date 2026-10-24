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
