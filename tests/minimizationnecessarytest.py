import time, win32gui, win32process, psutil
import win32ui
import dde
import os
from datetime import datetime

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

def find_book1_window_filtered():
    """Find Book1 Excel window, excluding captured files."""
    excel_windows = list_excel_windows()
    for pid, hwnd, title, vis in excel_windows:
        if vis and 'book1' in title.lower():
            if 'captured_' not in title.lower():
                return pid, hwnd, title
    return None, None, None

def minimize_other_excel_windows(target_hwnd):
    """Minimize all Excel windows except target."""
    excel_windows = list_excel_windows()
    minimized_windows = []
    
    for pid, hwnd, title, vis in excel_windows:
        if vis and hwnd != target_hwnd:
            try:
                exe = psutil.Process(pid).name()
                if exe.lower() == "excel.exe":
                    original_state = win32gui.GetWindowPlacement(hwnd)
                    result = win32gui.ShowWindow(hwnd, 6)  # SW_MINIMIZE
                    if result:
                        minimized_windows.append((hwnd, title, original_state))
            except:
                continue
    
    time.sleep(1)
    return minimized_windows

def restore_minimized_windows(minimized_windows):
    """Restore previously minimized Excel windows."""
    for hwnd, title, original_state in minimized_windows:
        try:
            win32gui.SetWindowPlacement(hwnd, original_state)
        except:
            pass

def bring_to_foreground(target_hwnd):
    """Bring target window to foreground."""
    win32gui.ShowWindow(target_hwnd, 9)  # SW_RESTORE
    time.sleep(0.5)
    win32gui.SetForegroundWindow(target_hwnd)
    time.sleep(0.5)
    win32gui.BringWindowToTop(target_hwnd)
    time.sleep(0.5)

def test_dde_save(test_name, target_title):
    """Test DDE save operation and return result."""
    test_path = f"C:\\temp\\{test_name}_{datetime.now().strftime('%H%M%S')}.xlsx"
    
    # Clean up existing file
    if os.path.exists(test_path):
        os.remove(test_path)
    
    try:
        # Create DDE server
        server = dde.CreateServer()
        server.Create("TestClient")
        
        # Connect to Excel
        conversation = dde.CreateConversation(server)
        conversation.ConnectTo("Excel", "System")
        
        # Simple save operation
        result = conversation.Exec(f'[SAVE.AS("{test_path}")]')
        
        # Wait and check result
        time.sleep(2)
        
        success = False
        if os.path.exists(test_path):
            file_size = os.path.getsize(test_path)
            if file_size > 0:
                success = True
                # Check what workbook was actually saved by looking at the window title
                excel_windows = list_excel_windows()
                saved_from_title = "Unknown"
                for pid, hwnd, title, vis in excel_windows:
                    if vis and psutil.Process(pid).name().lower() == "excel.exe":
                        if test_name.lower() in title.lower():
                            saved_from_title = title
                            break
            
            # Clean up test file
            try:
                os.remove(test_path)
            except:
                pass
        
        # Close DDE
        try:
            conversation.Close()
        except:
            pass
        server.Shutdown()
        
        return success, saved_from_title if success else "Failed", file_size if success else 0
        
    except Exception as e:
        return False, f"Error: {e}", 0

def show_excel_state():
    """Show current Excel window state."""
    excel_windows = list_excel_windows()
    
    print("\nCurrent Excel Windows:")
    print(f"{'Order':>5}  {'PID':>7}  {'Visible':^7}  {'Title'}")
    print("─" * 60)
    
    for i, (pid, hwnd, title, vis) in enumerate(excel_windows):
        exe = psutil.Process(pid).name()
        if exe.lower() == "excel.exe":
            print(f"{i:>5}  {pid:>7}  {str(vis):^7}  {title}")
    
    # Show foreground window
    try:
        fg_hwnd = win32gui.GetForegroundWindow()
        fg_title = win32gui.GetWindowText(fg_hwnd)
        print(f"\nForeground Window: {fg_title}")
    except:
        print(f"\nForeground Window: Could not determine")

def main():
    print("=" * 70)
    print("MINIMIZATION NECESSITY TEST")
    print("=" * 70)
    print("This test compares foreground-only vs minimization+foreground approaches")
    print("Make sure Book1 is pulsing orange and you have other Excel files open!")
    print("=" * 70)
    
    # Find target Book1
    target_result = find_book1_window_filtered()
    if not target_result[0]:
        print("\n❌ No Book1 found! Test cannot continue.")
        return
    
    target_pid, target_hwnd, target_title = target_result
    print(f"\n🎯 Target Book1:")
    print(f"   PID: {target_pid}")
    print(f"   Title: {target_title}")
    
    show_excel_state()
    
    print(f"\n" + "="*70)
    print("TEST 1: FOREGROUND ONLY (NO MINIMIZATION)")
    print("="*70)
    
    # Test 1: Foreground only
    print("🎯 Bringing Book1 to foreground...")
    bring_to_foreground(target_hwnd)
    
    print("💾 Testing DDE save (foreground only)...")
    success1, saved_from1, size1 = test_dde_save("foreground_only", target_title)
    
    print(f"Result: {'✅ SUCCESS' if success1 else '❌ FAILED'}")
    print(f"Saved from: {saved_from1}")
    print(f"File size: {size1} bytes")
    
    show_excel_state()
    
    # Wait a moment between tests
    time.sleep(2)
    
    print(f"\n" + "="*70)
    print("TEST 2: MINIMIZATION + FOREGROUND (FULL APPROACH)")
    print("="*70)
    
    # Test 2: Full approach with minimization
    print("🔽 Minimizing other Excel windows...")
    minimized_windows = minimize_other_excel_windows(target_hwnd)
    print(f"   Minimized {len(minimized_windows)} windows")
    
    print("🎯 Bringing Book1 to foreground...")
    bring_to_foreground(target_hwnd)
    
    print("💾 Testing DDE save (full approach)...")
    success2, saved_from2, size2 = test_dde_save("full_approach", target_title)
    
    print(f"Result: {'✅ SUCCESS' if success2 else '❌ FAILED'}")
    print(f"Saved from: {saved_from2}")
    print(f"File size: {size2} bytes")
    
    # Restore windows
    print("🔼 Restoring minimized windows...")
    restore_minimized_windows(minimized_windows)
    
    show_excel_state()
    
    print(f"\n" + "="*70)
    print("ANALYSIS & RECOMMENDATION")
    print("="*70)
    
    print(f"Test 1 (Foreground Only): {'✅ SUCCESS' if success1 else '❌ FAILED'}")
    print(f"Test 2 (Minimization + Foreground): {'✅ SUCCESS' if success2 else '❌ FAILED'}")
    
    if success1 and success2:
        print(f"\n🎯 CONCLUSION: Both approaches work!")
        print(f"📊 Comparison:")
        print(f"   Foreground Only: {size1} bytes from '{saved_from1}'")
        print(f"   Full Approach: {size2} bytes from '{saved_from2}'")
        
        if "Book1" in saved_from1 and "Book1" in saved_from2:
            print(f"\n✅ RECOMMENDATION: Use FOREGROUND ONLY approach")
            print(f"   - Simpler code")
            print(f"   - Less disruptive to user")
            print(f"   - Same targeting accuracy")
        else:
            print(f"\n⚠️ RECOMMENDATION: Keep FULL APPROACH")
            print(f"   - One of the methods targeted wrong workbook")
            
    elif success1 and not success2:
        print(f"\n🤔 UNEXPECTED: Foreground only works, full approach doesn't")
        print(f"✅ RECOMMENDATION: Use FOREGROUND ONLY")
        
    elif not success1 and success2:
        print(f"\n📊 EXPECTED: Full approach more reliable")
        print(f"✅ RECOMMENDATION: Keep FULL APPROACH with minimization")
        
    else:
        print(f"\n❌ BOTH FAILED: Need to investigate further")
        print(f"🔍 Check if Book1 is actually the target workbook")
    
    print(f"\n" + "="*70)

if __name__ == "__main__":
    main()