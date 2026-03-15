import os
import win32gui
import win32con
import win32api
import pythoncom
from win32com.shell import shell
from PIL import Image
from datetime import datetime
import threading

# Prevent DecompressionBombError for massive images
Image.MAX_IMAGE_PIXELS = None

# Custom Win32 User Messages for cross-thread communication
WM_MERGE_DONE = win32con.WM_USER + 1
WM_MERGE_ERROR = win32con.WM_USER + 2

PBM_SETRANGE = 0x0401
PBM_SETPOS = 0x0402
BIF_RETURNONLYFSDIRS = 0x0001
BIF_NEWDIALOGSTYLE = 0x0040

try:
    RESAMPLE_FILTER = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE_FILTER = Image.LANCZOS

# ==========================================
# MODULE 1: CORE IMAGE PROCESSING ENGINE
# Completely decoupled from GUI logic
# ==========================================
class ImageProcessor:
    
    @staticmethod
    def process(files, out_path, is_vertical, is_largest, progress_callback):
        if not files:
            raise ValueError("No files to process")

        sizes = []
        for p in files:
            with Image.open(p) as im:
                sizes.append((im.width, im.height))

        if is_vertical:
            # Determine target width based on user selection
            target_w = max(w for w, h in sizes) if is_largest else min(w for w, h in sizes)
            total_h = sum(int(h * target_w / w) if w != target_w else h for w, h in sizes)
            canvas = Image.new("RGB", (target_w, total_h))
            y = 0
            
            for i, p in enumerate(files):
                with Image.open(p) as orig_im:
                    im = orig_im.convert("RGB")
                    if im.width != target_w:
                        r = target_w / im.width
                        resized = im.resize((target_w, int(im.height * r)), RESAMPLE_FILTER)
                        im.close()
                        im = resized
                        
                    canvas.paste(im, (0, y))
                    y += im.height
                    im.close() # Free memory immediately
                    
                progress_callback(i + 1)
                
            canvas.save(out_path)
            canvas.close()
            
        else:
            # Determine target height based on user selection
            target_h = max(h for w, h in sizes) if is_largest else min(h for w, h in sizes)
            total_w = sum(int(w * target_h / h) if h != target_h else w for w, h in sizes)
            canvas = Image.new("RGB", (total_w, target_h))
            x = 0
            
            for i, p in enumerate(files):
                with Image.open(p) as orig_im:
                    im = orig_im.convert("RGB")
                    if im.height != target_h:
                        r = target_h / im.height
                        resized = im.resize((int(im.width * r), target_h), RESAMPLE_FILTER)
                        im.close()
                        im = resized
                        
                    canvas.paste(im, (x, 0))
                    x += im.width
                    im.close() # Free memory immediately
                    
                progress_callback(i + 1)
                
            canvas.save(out_path)
            canvas.close()

# ==========================================
# MODULE 2: WIN32 GUI & EVENT LOOP
# Handles purely visual elements and user inputs
# ==========================================
class AppGUI:
    def __init__(self):
        # Init COM to prevent modern folder browser crashes
        pythoncom.CoInitialize()
        win32gui.InitCommonControls()
        self.hinst = win32api.GetModuleHandle(None)
        
        self.output_dir = os.getcwd()
        self.files = []
        self.merge_msg = ""
        
        wc = win32gui.WNDCLASS()
        wc.hInstance = self.hinst
        wc.lpszClassName = "ImgMerger"
        wc.lpfnWndProc = self.WndProc
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        wc.hbrBackground = win32con.COLOR_WINDOW + 1

        try:
            win32gui.RegisterClass(wc)
        except:
            pass

        self.hwnd = win32gui.CreateWindow(
            "ImgMerger", "kishiMerger Image",
            win32con.WS_OVERLAPPED | win32con.WS_CAPTION |
            win32con.WS_SYSMENU | win32con.WS_MINIMIZEBOX |
            win32con.WS_VISIBLE,
            200, 200, 640, 520, 
            0, 0, self.hinst, None
        )

        lf = win32gui.LOGFONT()
        lf.lfFaceName = "Segoe UI"
        lf.lfHeight = -16
        lf.lfQuality = win32con.CLEARTYPE_QUALITY
        self.font = win32gui.CreateFontIndirect(lf)

        # UI Group: Direction
        win32gui.CreateWindow("STATIC", "Direction:", win32con.WS_CHILD | win32con.WS_VISIBLE, 10, 10, 80, 20, self.hwnd, 0, self.hinst, None)
        self.rV = win32gui.CreateWindow("BUTTON", "Vertical", win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_AUTORADIOBUTTON | win32con.WS_GROUP, 100, 10, 100, 20, self.hwnd, 101, self.hinst, None)
        self.rH = win32gui.CreateWindow("BUTTON", "Horizontal", win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_AUTORADIOBUTTON, 210, 10, 120, 20, self.hwnd, 102, self.hinst, None)
        win32gui.SendMessage(self.rV, win32con.BM_SETCHECK, win32con.BST_CHECKED, 0)

        # UI Group: File List & Buttons
        self.lb = win32gui.CreateWindowEx(win32con.WS_EX_CLIENTEDGE, "LISTBOX", "", win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_VSCROLL | win32con.WS_HSCROLL | win32con.LBS_NOTIFY | win32con.LBS_HASSTRINGS, 10, 50, 460, 220, self.hwnd, 103, self.hinst, None)
        self.btnAdd = win32gui.CreateWindow("BUTTON", "Add Files", win32con.WS_CHILD | win32con.WS_VISIBLE, 480, 50, 140, 30, self.hwnd, 104, self.hinst, None)
        self.btnUp = win32gui.CreateWindow("BUTTON", "Move Up", win32con.WS_CHILD | win32con.WS_VISIBLE, 480, 90, 140, 30, self.hwnd, 105, self.hinst, None)
        self.btnDown = win32gui.CreateWindow("BUTTON", "Move Down", win32con.WS_CHILD | win32con.WS_VISIBLE, 480, 130, 140, 30, self.hwnd, 106, self.hinst, None)
        self.btnRem = win32gui.CreateWindow("BUTTON", "Remove", win32con.WS_CHILD | win32con.WS_VISIBLE, 480, 170, 140, 30, self.hwnd, 107, self.hinst, None)
        
        # UI Group: Resize Mode (New feature)
        win32gui.CreateWindow("STATIC", "Resize base on:", win32con.WS_CHILD | win32con.WS_VISIBLE, 10, 280, 120, 20, self.hwnd, 0, self.hinst, None)
        self.rLargest = win32gui.CreateWindow("BUTTON", "Largest (Default)", win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_AUTORADIOBUTTON | win32con.WS_GROUP, 130, 280, 150, 20, self.hwnd, 111, self.hinst, None)
        self.rSmallest = win32gui.CreateWindow("BUTTON", "Smallest", win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_AUTORADIOBUTTON, 290, 280, 100, 20, self.hwnd, 112, self.hinst, None)
        win32gui.SendMessage(self.rLargest, win32con.BM_SETCHECK, win32con.BST_CHECKED, 0)

        # UI Group: Output Settings
        win32gui.CreateWindow("STATIC", "Output file name:", win32con.WS_CHILD | win32con.WS_VISIBLE, 10, 310, 150, 20, self.hwnd, 0, self.hinst, None)
        self.outEdit = win32gui.CreateWindowEx(win32con.WS_EX_CLIENTEDGE, "EDIT", "", win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.ES_AUTOHSCROLL, 10, 335, 300, 24, self.hwnd, 108, self.hinst, None)
        win32gui.SendMessage(self.outEdit, 0x1501, True, "File name (empty = auto timestamp)")
        self.btnFolder = win32gui.CreateWindow("BUTTON", "Browse Folder", win32con.WS_CHILD | win32con.WS_VISIBLE, 320, 335, 150, 24, self.hwnd, 110, self.hinst, None)
        self.outDirPath = win32gui.CreateWindow("STATIC", "Output: " + self.output_dir, win32con.WS_CHILD | win32con.WS_VISIBLE, 10, 365, 610, 20, self.hwnd, 0, self.hinst, None)
        
        # UI Group: Action & Progress
        self.btnMerge = win32gui.CreateWindow("BUTTON", "Merge Images", win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_DEFPUSHBUTTON, 10, 390, 610, 40, self.hwnd, 109, self.hinst, None)
        self.progress = win32gui.CreateWindowEx(0, "msctls_progress32", "", win32con.WS_CHILD | win32con.WS_VISIBLE, 10, 440, 610, 20, self.hwnd, 200, self.hinst, None)

        controls_to_font = [self.rV, self.rH, self.lb, self.btnAdd, self.btnUp, self.btnDown, self.btnRem, self.outEdit, self.btnMerge, self.btnFolder, self.outDirPath, self.rLargest, self.rSmallest]
        for h in controls_to_font:
            win32gui.SendMessage(h, win32con.WM_SETFONT, self.font, True)

        win32gui.DragAcceptFiles(self.hwnd, True)

    def run(self):
        win32gui.PumpMessages()

    def refresh(self):
        win32gui.SendMessage(self.lb, win32con.LB_RESETCONTENT, 0, 0)
        maxw = 0
        hdc = win32gui.GetDC(self.lb)
        old_font = win32gui.SelectObject(hdc, self.font)
        for f in self.files:
            win32gui.SendMessage(self.lb, win32con.LB_ADDSTRING, 0, f)
            w, _ = win32gui.GetTextExtentPoint32(hdc, f)
            maxw = max(maxw, w)
        win32gui.SelectObject(hdc, old_font)
        win32gui.ReleaseDC(self.lb, hdc)
        win32gui.SendMessage(self.lb, win32con.LB_SETHORIZONTALEXTENT, maxw + 20, 0)

    def WndProc(self, hwnd, msg, wParam, lParam):
        if msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0
            
        # Callbacks from Worker Thread
        elif msg == WM_MERGE_DONE:
            win32gui.SendMessage(self.progress, PBM_SETPOS, 0, 0)
            win32gui.EnableWindow(self.btnMerge, True)
            win32gui.MessageBox(self.hwnd, self.merge_msg, "Success", win32con.MB_OK | win32con.MB_ICONINFORMATION)
            return 0
            
        elif msg == WM_MERGE_ERROR:
            win32gui.SendMessage(self.progress, PBM_SETPOS, 0, 0)
            win32gui.EnableWindow(self.btnMerge, True)
            win32gui.MessageBox(self.hwnd, self.merge_msg, "Error", win32con.MB_OK | win32con.MB_ICONERROR)
            return 0

        elif msg == win32con.WM_DROPFILES:
            count = win32api.DragQueryFile(wParam, -1)
            for i in range(count):
                f = win32api.DragQueryFile(wParam, i)
                if os.path.isfile(f):
                    self.files.append(f)
            win32api.DragFinish(wParam)
            self.files.sort(key=lambda x: os.path.basename(x).lower())
            self.refresh()

        elif msg == win32con.WM_COMMAND:
            cid = win32api.LOWORD(wParam)
            if cid == 104: self.add_files()
            elif cid == 105: self.move(-1)
            elif cid == 106: self.move(1)
            elif cid == 107: self.remove()
            elif cid == 109: self.start_merge_thread()
            elif cid == 110: self.pick_folder()

        return win32gui.DefWindowProc(hwnd, msg, wParam, lParam)

    def add_files(self):
        try:
            fname, _, _ = win32gui.GetOpenFileNameW(
                hwndOwner=self.hwnd,
                Flags=win32con.OFN_ALLOWMULTISELECT | win32con.OFN_EXPLORER | win32con.OFN_FILEMUSTEXIST,
                Filter="Images\0*.png;*.jpg;*.jpeg;*.webp\0\0",
                MaxFile=65536
            )
            parts = [p for p in fname.split("\0") if p]
            if len(parts) == 1:
                self.files.append(parts[0])
            elif len(parts) > 1:
                base = parts[0]
                for p in parts[1:]:
                    self.files.append(os.path.join(base, p))
            self.files.sort(key=lambda x: os.path.basename(x).lower())
            self.refresh()
        except Exception as e:
            if getattr(e, 'winerror', 0) != 0 and e.args[0] != 0:
                print(f"Error Add Files: {e}")

    def move(self, d):
        idx = win32gui.SendMessage(self.lb, win32con.LB_GETCURSEL, 0, 0)
        if idx < 0: return
        ni = idx + d
        if 0 <= ni < len(self.files):
            self.files[idx], self.files[ni] = self.files[ni], self.files[idx]
            self.refresh()
            win32gui.SendMessage(self.lb, win32con.LB_SETCURSEL, ni, 0)

    def remove(self):
        idx = win32gui.SendMessage(self.lb, win32con.LB_GETCURSEL, 0, 0)
        if idx >= 0:
            del self.files[idx]
            self.refresh()

    def pick_folder(self):
        try:
            pidl, _, _ = shell.SHBrowseForFolder(
                self.hwnd,
                None,
                "Select output folder",
                BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE,
                None,
                None
            )
            if pidl:
                path = shell.SHGetPathFromIDList(pidl)
                if path:
                    if isinstance(path, bytes):
                        path = path.decode('utf-8', errors='ignore')
                    self.output_dir = path
                    win32gui.SetWindowText(self.outDirPath, "Output: " + self.output_dir)
        except Exception as e:
            print(f"Error Browse Folder: {e}")
            win32gui.MessageBox(self.hwnd, f"Error: {str(e)}", "Error", win32con.MB_OK | win32con.MB_ICONERROR)

    def start_merge_thread(self):
        if len(self.files) < 2:
            win32gui.MessageBox(self.hwnd, "Minimum 2 images required", "Warning", win32con.MB_OK | win32con.MB_ICONWARNING)
            return
            
        name = win32gui.GetWindowText(self.outEdit).strip()
        if not name:
            name = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        if not name.endswith(".png"):
            name += ".png"
            
        out = os.path.join(self.output_dir, name)
        
        # Capture UI states before passing to thread
        vertical = win32gui.SendMessage(self.rV, win32con.BM_GETCHECK, 0, 0) == win32con.BST_CHECKED
        largest = win32gui.SendMessage(self.rLargest, win32con.BM_GETCHECK, 0, 0) == win32con.BST_CHECKED
        
        win32gui.SendMessage(self.progress, PBM_SETRANGE, 0, win32api.MAKELONG(0, len(self.files)))
        win32gui.SendMessage(self.progress, PBM_SETPOS, 0, 0)
        win32gui.EnableWindow(self.btnMerge, False)
        
        # Offload intensive task to worker thread
        threading.Thread(target=self._merge_worker, args=(vertical, largest, out), daemon=True).start()

    def _merge_worker(self, vertical, largest, out):
        try:
            # Inject Processor Module
            ImageProcessor.process(
                self.files, 
                out, 
                vertical, 
                largest, 
                lambda pos: win32gui.PostMessage(self.progress, PBM_SETPOS, pos, 0) # Progress callback
            )
            self.merge_msg = "Saved:\n" + out
            win32gui.PostMessage(self.hwnd, WM_MERGE_DONE, 0, 0)
        except Exception as e:
            self.merge_msg = f"Process failed:\n{str(e)}"
            win32gui.PostMessage(self.hwnd, WM_MERGE_ERROR, 0, 0)

if __name__ == "__main__":
    AppGUI().run()
