import time, win32gui, win32process, psutil

def list_excel_windows():
    """Return [(pid, hwnd, title, visible)] for every XLMAIN window."""
    windows = []
    def _enum(hwnd, _):
        if win32gui.GetClassName(hwnd) == "XLMAIN":
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            title  = win32gui.GetWindowText(hwnd)
            vis    = win32gui.IsWindowVisible(hwnd)
            windows.append((pid, hwnd, title, vis))
    win32gui.EnumWindows(_enum, None)
    return windows

# ---------- pretty-print once ----------
print(f"{'PID':>7}  {'HWND':>10}  {'Visible':^7}   Window title")
print("─"*60)
for pid, hwnd, title, vis in list_excel_windows():
    exe = psutil.Process(pid).name()
    if exe.lower() != "excel.exe":       # should always be Excel, but filter just in case
        continue
    print(f"{pid:>7}  {hwnd:#010x}   {str(vis):^7}   {title}")

# ---------- or watch continuously ----------
# while True:
#     os.system("cls")
#     ...
#     time.sleep(1)
