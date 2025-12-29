---
title: ComfyUI Model Management
description: PyQt6 desktop GUI to scan, sort, export, and delete ComfyUI model files
---

# ComfyUI Model Management

A polished **Windows desktop GUI** built with **PyQt6** to help you manage ComfyUI model files.

The application scans your ComfyUI directory (and all subdirectories) for common model formats, displays them in a sortable and filterable grid, and allows you to safely delete or export results to Excel.

---

## Screenshots

![Dark Mode](https://raw.githubusercontent.com/Rhovanx/comfyui-model-management/main/assets/screenshots/app-dark.png)
![Light Mode](https://raw.githubusercontent.com/Rhovanx/comfyui-model-management/main/assets/screenshots/app-light.png)

---

## Features

- Scan folders recursively for:
  - `.safetensors`, `.ckpt`, `.pth`, `.pt`, `.onnx`, `.bin`, `.gguf`
- Frozen checkbox column (always visible while scrolling horizontally)
- Sort by any column with ▲ / ▼ indicator  
  - Default: **LastAccessTime (ascending)** to surface least-used models first
- Filter/search by model name, directory, or extension
- Select All / Select None
- Delete selected models:
  - ✅ Move to Recycle Bin (default, safe)
  - ✅ Permanent delete (optional)
- Export visible results to Excel (`.xlsx`)
  - Automatically opens Excel if available
- Context-aware progress bar:
  - Scan progress
  - Delete progress
  - Selection summary (count + total size)
- Light / Dark theme switch with high-contrast UI
- Remembers user settings:
  - Last ComfyUI folder
  - Theme (Light/Dark)
  - Last sorted column and direction

---

## Quick Start

### Clone the repository

```bash
git clone https://github.com/Rhovanx/comfyui-model-management.git
cd comfyui-model-management
```

### (Optional) Use a virtual environment

```bash
python -m venv .venv
.\.venv\Scripts\activate
```

### Install dependencies

```bash
pip install -r requirements.txt
```

---

## Usage

Run the application:

```bash
python src/comfyui_model_management.py
```

---

## Notes

- **Recycle Bin deletion** requires:
  - `send2trash`
- **Open Excel after export** requires:
  - Microsoft Excel installed
  - `pywin32`

If Excel is not available, the export still succeeds and the application will simply skip launching Excel.

---

## Build a Windows `.exe` (PyInstaller)

### 1) Install PyInstaller

```bash
pip install pyinstaller
```

### 2) Build a one-file executable

```bash
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "ComfyUI-Model-Management" ^
  src\comfyui_model_management.py
```

The executable will be located at:

```
dist\ComfyUI-Model-Management.exe
```

---

### 3) Recommended build (folder-based, faster startup)

```bash
pyinstaller --noconfirm --clean --windowed ^
  --name "ComfyUI-Model-Management" ^
  src\comfyui_model_management.py
```

This produces:

```
dist\ComfyUI-Model-Management\
```

---

### 4) PyQt6 troubleshooting (if needed)

If you encounter missing Qt plugin errors, rebuild with:

```bash
pyinstaller --noconfirm --clean --onefile --windowed ^
  --collect-all PyQt6 ^
  src\comfyui_model_management.py
```

---

### 5) Optional: application icon

Place an icon at:

```
assets\icon.ico
```

Then build with:

```bash
pyinstaller --noconfirm --clean --onefile --windowed ^
  --icon assets\icon.ico ^
  --name "ComfyUI-Model-Management" ^
  src\comfyui_model_management.py
```

---

## License

MIT
