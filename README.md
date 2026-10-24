# WhatsApp Messages Sender

A GUI + CLI tool for sending WhatsApp messages (with or without images) to contacts stored in an Excel file. Built with Python, `customtkinter`, and `pywhatkit`.

## Features

- **Send Messages** - Send personalized messages to contacts from `data.xlsx`
- **Send with Images** - Attach an image from the `images/` folder
- **Contact Management (GUI)** - Add, edit, delete, and preview contacts with Arabic (RTL) text support
- **Phone Validation** - Validates phone numbers and normalizes them (removes leading 0)
- **Duplicate Detection** - Find and remove duplicate contacts in the preview window
- **Sent/Unsent Tracking** - Move contacts between "Data" and "Sent" sheets
- **Pagination** - Efficient browsing of large contact lists
- **Bilingual Support** - Proper Arabic text shaping and display

## Requirements

- Python 3.x
- Dependencies: `customtkinter`, `openpyxl`, `pywhatkit`, `arabic_reshaper`, `python-bidi`, `rich`

Install dependencies:
```bash
pip install customtkinter openpyxl pywhatkit arabic-reshaper python-bidi rich
```

## Configuration

Edit `config.ini` to customize behavior:

```ini
[settings]
load_data = data.xlsx   # Contacts workbook
wait_time = 10          # Seconds to wait before sending
tab_close = True        # Close browser tab after sending
image_folder = images/  # Folder containing images to send
page_size = 25          # Rows per page in GUI preview
```

## Usage

### CLI Mode
```bash
python main.py
```

Menu options:
1. Send message
2. Send message with image
3. Add / Edit Contacts (opens GUI)
4. About
0. Exit

### GUI Mode
```bash
python main.py --gui
# or
python gui.py
```

Features in GUI:
- **Add Contact** - Add new contacts with name and phone number
- **Preview Data** - View, edit, delete contacts with pagination
- **Delete Duplicates** - Remove duplicate phone numbers
- **Sent View** - View and restore sent contacts
- **Arabic Support** - Type and display Arabic text correctly


## Notes

- Phone numbers are normalized (leading `0` removed) and validated
- Contacts are stored in Excel format with `data` and `sent`/`finished` sheets
- For WhatsApp Web automation, `pywhatkit` will open a browser tab
- On Linux, `copyq` may be used for image copying

## License

This project is for personal use.
