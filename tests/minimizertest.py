import time, win32gui, win32process, psutil
import win32ui
import dde
import os

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
    """Filtered detection logic - exclude captured files."""
    excel_windows = list_excel_windows()
    for pid, hwnd, title, vis in excel_windows:
        if vis and 'book1' in title.lower():
            if 'captured_' not in title.lower():
                return pid, hwnd, title
    return None, None, None

def show_excel_state(stage_name):
    """Show current Excel window state."""
    print(f"\n{'='*60}")
    print(f"STAGE: {stage_name}")
    print(f"{'='*60}")
    
    excel_windows = list_excel_windows()
    
    print("All Excel Windows:")
    print(f"{'Order':>5}  {'PID':>7}  {'HWND':>10}  {'Vis':^3}  {'Title'}")
    print("─" * 70)
    
    book1_windows = []
    for i, (pid, hwnd, title, vis) in enumerate(excel_windows):
        exe = psutil.Process(pid).name()
        if exe.lower() == "excel.exe":
            print(f"{i:>5}  {pid:>7}  {hwnd:#010x}  {str(vis):^3}  {title}")
            if 'book1' in title.lower():
                book1_windows.append((i, pid, hwnd, title, vis))
    
    # Show foreground window
    try:
        fg_hwnd = win32gui.GetForegroundWindow()
        fg_title = win32gui.GetWindowText(fg_hwnd)
        print(f"\nForeground Window: {fg_hwnd:#010x} - {fg_title}")
    except:
        print(f"\nForeground Window: Could not determine")
    
    return excel_windows, book1_windows

def minimize_other_excel_windows_test(target_hwnd):
    """Test version of minimize function."""
    print(f"\n🔽 MINIMIZING all Excel windows except target {target_hwnd:#010x}...")
    
    excel_windows = list_excel_windows()
    minimized_windows = []
    
    for pid, hwnd, title, vis in excel_windows:
        if vis and hwnd != target_hwnd:
            try:
                exe = psutil.Process(pid).name()
                if exe.lower() == "excel.exe":
                    print(f"   Minimizing: {title}")
                    original_state = win32gui.GetWindowPlacement(hwnd)
                    result = win32gui.ShowWindow(hwnd, 6)  # SW_MINIMIZE
                    if result:
                        minimized_windows.append((hwnd, title, original_state))
                        print(f"     ✅ Minimized successfully")
            except Exception as e:
                print(f"     ❌ Error minimizing: {e}")
                continue
    
    print(f"   Total minimized: {len(minimized_windows)} Excel windows")
    time.sleep(1)
    return minimized_windows

def bring_to_foreground_test(target_hwnd, target_title):
    """Test bringing target window to foreground."""
    print(f"\n🎯 BRINGING TARGET TO FOREGROUND...")
    print(f"   Target: {target_hwnd:#010x} - {target_title}")
    
    try:
        # Multiple methods to ensure window becomes foreground
        print(f"   Step 1: ShowWindow(SW_RESTORE)")
        win32gui.ShowWindow(target_hwnd, 9)  # SW_RESTORE
        time.sleep(0.5)
        
        print(f"   Step 2: SetForegroundWindow")
        win32gui.SetForegroundWindow(target_hwnd)
        time.sleep(0.5)
        
        print(f"   Step 3: BringWindowToTop")
        win32gui.BringWindowToTop(target_hwnd)
        time.sleep(0.5)
        
        # Verify it worked
        current_fg = win32gui.GetForegroundWindow()
        if current_fg == target_hwnd:
            print(f"   ✅ Successfully brought target to foreground")
        else:
            current_title = win32gui.GetWindowText(current_fg)
            print(f"   ⚠️ Foreground is different window: {current_fg:#010x} - {current_title}")
            
    except Exception as e:
        print(f"   ❌ Error bringing to foreground: {e}")

def test_dde_connection_detailed():
    """Test DDE connection with detailed analysis."""
    print(f"\n🔗 TESTING DDE CONNECTION (DETAILED)...")
    
    try:
        # Create DDE server
        server = dde.CreateServer()
        server.Create("TestClient")
        
        # Connect to Excel
        conversation = dde.CreateConversation(server)
        conversation.ConnectTo("Excel", "System")
        
        print(f"   ✅ DDE Connected successfully")
        
        # Test 1: Try to get info about connected Excel
        try:
            print(f"   Test 1: Getting Excel info...")
            # This command should tell us something about the connected Excel
            result = conversation.Exec('[ECHO("DDE_CONNECTED")]')
            print(f"     ECHO result: {result}")
        except Exception as e:
            print(f"     ECHO failed: {e}")
        
        # Test 2: Try a simple operation to see which Excel responds
        try:
            print(f"   Test 2: Testing window activation...")
            # This should maximize the Excel window that DDE is connected to
            result = conversation.Exec('[APP.MAXIMIZE()]')
            print(f"     APP.MAXIMIZE result: {result}")
            time.sleep(1)
        except Exception as e:
            print(f"     APP.MAXIMIZE failed: {e}")
        
        # Test 3: Try to save a test file to see which workbook responds
        test_path = "C:\\temp\\dde_test.xlsx"
        try:
            print(f"   Test 3: Testing save operation...")
            # Clean up any existing test file
            if os.path.exists(test_path):
                os.remove(test_path)
            
            result = conversation.Exec(f'[SAVE.AS("{test_path}")]')
            print(f"     SAVE.AS result: {result}")
            
            time.sleep(2)
            if os.path.exists(test_path):
                print(f"     ✅ Test file created: {test_path}")
                try:
                    os.remove(test_path)
                    print(f"     🗑️ Test file cleaned up")
                except:
                    pass
            else:
                print(f"     ⚠️ Test file not created")
                
        except Exception as e:
            print(f"     SAVE test failed: {e}")
        
        # Close connection
        try:
            conversation.Close()
        except:
            pass
        server.Shutdown()
        
        print(f"   ✅ DDE connection test completed")
        
    except Exception as e:
        print(f"   ❌ DDE connection failed: {e}")

def restore_minimized_windows_test(minimized_windows):
    """Restore previously minimized Excel windows."""
    if not minimized_windows:
        return
    
    print(f"\n🔼 RESTORING {len(minimized_windows)} minimized Excel windows...")
    
    for hwnd, title, original_state in minimized_windows:
        try:
            win32gui.SetWindowPlacement(hwnd, original_state)
            print(f"   ✅ Restored: {title}")
        except Exception as e:
            print(f"   ⚠️ Could not restore {title}: {e}")

def main():
    print("=" * 70)
    print("FOREGROUND WINDOW DDE TARGETING TEST")
    print("=" * 70)
    print("This test checks if bringing Book1 to foreground fixes DDE targeting")
    print("Run this while Book1 is pulsing orange!")
    print("=" * 70)
    
    # STAGE 1: Initial state
    show_excel_state("1. INITIAL STATE (Book1 pulsing orange)")
    
    # Find target using filtered logic
    target_result = find_book1_window_filtered()
    if not target_result[0]:
        print("\n❌ No target Book1 found! Test cannot continue.")
        return
    
    target_pid, target_hwnd, target_title = target_result
    print(f"\n🎯 TARGET SELECTED:")
    print(f"   PID: {target_pid}")
    print(f"   HWND: {target_hwnd:#010x}")
    print(f"   Title: {target_title}")
    
    # STAGE 2: Minimize other windows
    minimized_windows = minimize_other_excel_windows_test(target_hwnd)
    show_excel_state("2. AFTER MINIMIZATION")
    
    # STAGE 3: Bring target to foreground (NEW!)
    bring_to_foreground_test(target_hwnd, target_title)
    show_excel_state("3. AFTER BRINGING TARGET TO FOREGROUND")
    
    # STAGE 4: Test DDE connection
    test_dde_connection_detailed()
    show_excel_state("4. AFTER DDE TEST")
    
    # STAGE 5: Restore
    restore_minimized_windows_test(minimized_windows)
    show_excel_state("5. AFTER RESTORATION")
    
    print(f"\n" + "="*70)
    print("ANALYSIS")
    print("="*70)
    print("Check the DDE test results above:")
    print("- Did APP.MAXIMIZE affect the target Book1 window?")
    print("- Did the test file get saved from the target Book1?")
    print("- Did bringing target to foreground change DDE targeting?")
    print("="*70)

if __name__ == "__main__":
    main()