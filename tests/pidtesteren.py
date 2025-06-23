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

def detection_test():
    print("=== BOOK1 DETECTION TEST ===")
    print("Run this while Book1 is pulsating orange")
    print("\n")
    
    # Show all Excel windows first
    print("ALL Excel Windows:")
    print(f"{'PID':>7}  {'HWND':>10}  {'Visible':^7}   Window title")
    print("─" * 60)
    
    excel_windows = list_excel_windows()
    for i, (pid, hwnd, title, vis) in enumerate(excel_windows):
        exe = psutil.Process(pid).name()
        if exe.lower() == "excel.exe":
            print(f"{pid:>7}  {hwnd:#010x}   {str(vis):^7}   {title}")
    
    print("\n" + "=" * 60)
    
    # Now find all Book1 windows specifically
    book1_windows = []
    for i, (pid, hwnd, title, vis) in enumerate(excel_windows):
        exe = psutil.Process(pid).name()
        if exe.lower() == "excel.exe" and 'book1' in title.lower():
            book1_windows.append((i, pid, hwnd, title, vis))
    
    print("BOOK1 Windows Found:")
    if book1_windows:
        print(f"{'Order':>5}  {'PID':>7}  {'HWND':>10}  {'Visible':^7}   Window title")
        print("─" * 70)
        
        for order, pid, hwnd, title, vis in book1_windows:
            print(f"{order:>5}  {pid:>7}  {hwnd:#010x}   {str(vis):^7}   {title}")
        
        print(f"\n📊 SUMMARY:")
        print(f"   Total Book1 windows found: {len(book1_windows)}")
        
        # Show which one the current logic would pick
        first_book1 = book1_windows[0]
        print(f"   Current script would pick: Order {first_book1[0]} - {first_book1[3]}")
        
        # Show additional details about each Book1
        print(f"\n🔍 DETAILED ANALYSIS:")
        for order, pid, hwnd, title, vis in book1_windows:
            print(f"   Order {order}:")
            print(f"     PID: {pid}")
            print(f"     HWND: {hwnd:#010x}")
            print(f"     Title: '{title}'")
            print(f"     Visible: {vis}")
            
            # Try to get window state information
            try:
                placement = win32gui.GetWindowPlacement(hwnd)
                show_state = placement[1]  # SW_HIDE=0, SW_NORMAL=1, SW_MINIMIZE=2, SW_MAXIMIZE=3
                state_names = {0: "HIDDEN", 1: "NORMAL", 2: "MINIMIZED", 3: "MAXIMIZED"}
                print(f"     Window State: {state_names.get(show_state, f'UNKNOWN({show_state})')}")
            except Exception as e:
                print(f"     Window State: ERROR - {e}")
            
            # Check if window is in foreground
            try:
                foreground_hwnd = win32gui.GetForegroundWindow()
                is_foreground = (hwnd == foreground_hwnd)
                print(f"     Is Foreground: {is_foreground}")
            except:
                print(f"     Is Foreground: ERROR")
            
            print()
    
    else:
        print("❌ No Book1 windows found!")
        print("   This means the detection logic is working correctly")
        print("   The issue might be elsewhere")
    
    print("=" * 60)
    print("INSTRUCTIONS:")
    print("1. Run this script while Book1 is pulsating orange")
    print("2. Look at how many Book1 windows are found")
    print("3. Note which one is 'Order 0' (first found)")
    print("4. Click the pulsating Book1 window")
    print("5. Run this script again to see what changed")

if __name__ == "__main__":
    detection_test()