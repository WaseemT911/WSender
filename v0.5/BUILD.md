# Building a portable .exe (Windows 7 & 10)

The goal: **one `WhatsAppSender.exe`** that runs on any Windows PC with **no Python
and no libraries installed**.

PyInstaller is **not a cross-compiler** — you must **build on Windows** (it cannot be
built from Linux/macOS). This machine is Linux, so the actual `.exe` must be produced
on a Windows box.

---

## 1. Why Python 3.8

| Target OS | Highest Python that runs there |
|---|---|
| Windows 7 | **3.8.x** (last version supporting Win7) |
| Windows 10 | any current Python |

For the exe to run on **both**, build it with **Python 3.8**:

- Use the **32-bit** Python 3.8 installer for maximum compatibility (runs on 32- and
  64-bit Windows). Use 64-bit only if you don't care about 32-bit PCs.
- Download: https://www.python.org/downloads/release/python-3810/ (Windows installer)
- Note: PyInstaller itself must be an older version on 3.8 (e.g. `pyinstaller<6.11`)
  if the newest one refuses to install.

## 2. Build steps (on Windows)

```bat
:: in the v0.5 folder
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

build.bat            :: or run the pyinstaller command manually (see below)
```

Manual equivalent of `build.bat`:

```bat
pyinstaller --noconfirm --onefile --name "WhatsAppSender" --collect-all customtkinter main.py
```

`--collect-all customtkinter` is **required** so the theme JSON files and assets under
`customtkinter/assets` are bundled.

Result: **`dist\WhatsAppSender.exe`** (single file).

## 3. What to ship to the target PC

Copy these into a folder (e.g. `C:\WhatsAppSender\` or a USB stick):

```
WhatsAppSender.exe
config.ini          (optional — created automatically if missing)
message.txt         (optional — created automatically if missing)
data.xlsx           (optional — created automatically if missing)
images\             (the image to send, e.g. WLogo.png)
```

The program is **portable**: it always reads/writes `config.ini`, `data.xlsx`,
`message.txt` and `images\` **next to the exe**, so everything stays on the USB stick /
folder. Missing files are created automatically on first run.

## 4. Windows 7 caveat (important)

Python 3.8 bundles the **Universal C Runtime (UCRT)**. On a clean Windows 7 the exe may
fail with a missing-DLL error. Fix by installing the free Microsoft update on the
Windows 7 PC (one time):

- **KB2999226** (Update for Universal C Runtime in Windows)
- or install **vc_redist.x86.exe / vc_redist.x64.exe** (Visual C++ 2015 Redistributable)

Windows 10 already includes the UCRT — nothing needed there.

## 5. How to launch

- **CLI:** double-click `WhatsAppSender.exe` (opens the menu).
- **GUI directly:** create a shortcut to `WhatsAppSender.exe --gui`, or use menu
  option **[3] Add / Edit Contacts** which opens the GUI.

## 6. Runtime requirements on the target PC

- **Google Chrome** (or Firefox) — `pywhatkit` opens WhatsApp Web in the browser and
  the user must be logged into https://web.whatsapp.com.
- A working internet connection.
- An image in `images\` for the "send with image" option.

---

## Notes / limitations

- `pywhatkit` sends messages by automating WhatsApp Web in the browser, so it is **not
  fully unattended** — the browser must open and WhatsApp Web must be logged in. This
  is unchanged from previous versions.
- On Linux the app also uses **CopyQ** (`copyq --start-server`) — not needed on Windows.
- First launch is slightly slow because the one-file exe unpacks to a temp folder.
