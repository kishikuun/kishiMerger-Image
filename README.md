# 🖼️ kishiMerger Image

![Python3](https://img.shields.io/badge/Python-3.x-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows%20(Win32%20API)-lightgrey.svg)
![License](https://img.shields.io/badge/License-Copyright%202026-green.svg)

A lightweight, highly optimized image merging utility built with Python and the native Win32 API. Developed under **kishi Studios**, this tool is designed for memory efficiency, native Windows performance, and high-speed image processing.

## ✨ Key Features
* **Native Win32 GUI:** Built entirely on Windows native components for a lightning-fast, zero-bloat interface.
* **Smart Memory Optimization:** Sequential processing with immediate garbage collection to handle massive image batches without RAM spikes.
* **Optimized Binary:** Compressed with **UPX (Ultimate Packer for eXecutables)** for a minimal disk footprint and fast in-memory decompression.
* **Flexible Scaling Modes:** Support for Vertical/Horizontal merging with the ability to scale based on the **Largest** or **Smallest** image in the set.
* **Responsive Multi-threading:** Decoupled GUI and Processing threads to ensure the window remains active and provides real-time progress updates.

## 🛠️ Build Instructions
This project requires Python 3.x, `Pillow`, and `pywin32`. We use `PyInstaller` combined with `UPX` to create the smallest possible standalone executable.

### 1. Setup Environment
Initialize a clean virtual environment and install dependencies:
```bash
python -m venv env
env\Scripts\activate
pip install pyinstaller pillow pywin32
```

### 2. Prepare UPX (Compression)
To minimize the `.exe` size, download the [UPX](https://github.com/upx/upx/releases) and place `upx.exe` directly into the project root directory.

### 3. Compile the Application (Release Build)
Build the final executable using extreme optimization flags and UPX compression. Ensure your `version.txt` is present for metadata embedding:
```bash
pyinstaller --noconsole --onefile --version-file=version.txt --upx-dir . --exclude-module tkinter --exclude-module unittest --exclude-module email --exclude-module http --exclude-module pydoc --exclude-module xml main.py
```

🚀 Usage
1. Launch kishiMerger.exe.
2. Drag and drop images into the listbox or use the Add Files button.
3. Choose your preferred Direction (Vertical/Horizontal) and Resize Base (Largest/Smallest).
4. Select the output directory and click Merge Images.
5. The tool will process the images and notify you upon completion.

👤 Author: kishikuun
