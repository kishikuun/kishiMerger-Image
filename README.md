# 🖼️ kishiMerger Image

![Python3](https://img.shields.io/badge/Python-3.x-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20(Win32%20API)-lightgrey.svg)
![License](https://img.shields.io/badge/License-Copyright%202026-green.svg)

A lightweight, highly optimized image merging utility built with Python and the native Win32 API. Developed under **kishi Studios**, this tool is designed for memory efficiency, native Windows performance, and handling massive image processing tasks seamlessly.

## ✨ Key Features
* **Native Win32 GUI:** Built entirely on Windows native components using `pywin32` for a clean, lightning-fast interface.
* **Smart Memory Optimization:** Images are processed, resized, and pasted sequentially with immediate garbage collection to prevent RAM spikes.
* **Highly Customizable:** Choose between Vertical or Horizontal merging, and scale based on the Largest (Default) or Smallest image in the batch.
* **Advanced List Management:** Drag & drop support directly from Windows Explorer, auto-sorting by name, and manual reordering.
* **Lightweight & Stable:** Multi-threaded engine completely decouples the GUI event loop from the image processing core.

## 🛠️ Build Instructions
This project uses Python, `Pillow`, and `pywin32`. To compile it into a single, standalone portable `.exe` with embedded metadata, use `PyInstaller` within a clean virtual environment.

### 1. Setup Environment
Initialize a clean virtual environment and install the required dependencies:
```bash
python -m venv env
env\Scripts\activate
pip install pyinstaller pillow pywin32
```

### 2. Compile the Application (Release Build)
Build the final executable using PyInstaller. The --noconsole flag hides the background command prompt, and --onefile packs all dependencies into a single binary. Ensure your version.txt is in the root directory:
```bash
pyinstaller --noconsole --onefile --version-file=version.txt main.py
```
🚀 Usage
1. Launch kishiMerger.exe.
2. Drag and drop your images into the listbox or click Add Files.
(Optional but recommended) Rearrange the image order using Move Up or Move Down.
3. Select the Direction and Resize base according to your needs.
4. Click Merge Images to begin execution. The output will be saved to your selected output folder.

👤 Author: kishikuun
