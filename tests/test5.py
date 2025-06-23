import time, win32gui, win32process, psutil
import win32ui, win32con
import dde
import os
from datetime import datetime

# Target directory
SAVE_FOLDER = r"C:\Users\sasuk\Documents\CapturedExports"

def extract_workbook_name(window_title):
    """Extract the workbook name from Excel window title."""
    # Excel window titles are in format: "WorkbookName - Excel"
    if " - Excel" in window_title:
        return window_title.replace(" - Excel", "").strip()
    return "UnknownWorkbook"

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

def find_book1_window():
    """Find an unsaved Excel workbook window (Book1, Book2, etc.)."""
    excel_windows = list_excel_windows()
    
    for pid, hwnd, title, vis in excel_windows:
        if vis:
            workbook_name = extract_workbook_name(title)
            # Look for unsaved workbooks (Book1, Book2, Book3, etc.)
            # Avoid our own captured files (which start with "Captured_")
            if (workbook_name.lower().startswith('book') and 
                not workbook_name.lower().startswith('captured_')):
                return pid, hwnd, title
    return None, None, None

def minimize_other_excel_windows(target_hwnd):
    """Minimize all Excel windows except the target to make it primary."""
    
    print("🔽 Minimizing other Excel windows to make target primary...")
    
    excel_windows = list_excel_windows()
    minimized_windows = []
    
    for pid, hwnd, title, vis in excel_windows:
        if vis and hwnd != target_hwnd:  # Skip target window itself
            try:
                exe = psutil.Process(pid).name()
                if exe.lower() == "excel.exe":
                    print(f"   Minimizing: PID {pid} - {title}")
                    
                    # Store original state for potential restoration
                    original_state = win32gui.GetWindowPlacement(hwnd)
                    
                    # Minimize the window
                    result = win32gui.ShowWindow(hwnd, 6)  # SW_MINIMIZE
                    
                    if result:
                        minimized_windows.append((hwnd, title, original_state))
                        print(f"     ✅ Minimized successfully")
                    else:
                        print(f"     ⚠️ Minimize may have failed")
                        
            except Exception as e:
                print(f"     ❌ Error minimizing {hwnd:#010x}: {e}")
                continue
    
    print(f"   Total minimized: {len(minimized_windows)} Excel windows")
    
    # Give Windows time to process the minimize operations
    time.sleep(1)
    
    return minimized_windows

def restore_minimized_windows(minimized_windows):
    """Restore previously minimized Excel windows."""
    
    if not minimized_windows:
        return
    
    print(f"🔼 Restoring {len(minimized_windows)} minimized Excel windows...")
    
    for hwnd, title, original_state in minimized_windows:
        try:
            # Restore to original state
            win32gui.SetWindowPlacement(hwnd, original_state)
            print(f"   ✅ Restored: {title}")
        except Exception as e:
            print(f"   ⚠️ Could not restore {title}: {e}")
            # Fallback - just show the window
            try:
                win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
                print(f"     Fallback restore successful")
            except:
                print(f"     Fallback restore also failed")

def prepare_save_location(window_title):
    """Prepare and clean the save location to avoid file conflicts."""
    
    # Ensure save folder exists
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    
    # Extract workbook name and create new filename pattern
    workbook_name = extract_workbook_name(window_title)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Captured_{workbook_name}_{timestamp}.xlsx"
    full_path = os.path.join(SAVE_FOLDER, filename)
    
    print(f"📁 Preparing save location: {full_path}")
    print(f"   Original workbook: {workbook_name}")
    
    # Check if file already exists and remove it
    if os.path.exists(full_path):
        try:
            os.remove(full_path)
            print(f"   ✅ Removed existing file")
        except Exception as e:
            print(f"   ⚠️ Could not remove existing file: {e}")
            # Try alternative filename with microseconds
            alt_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:17]
            filename = f"Captured_{workbook_name}_{alt_timestamp}.xlsx"
            full_path = os.path.join(SAVE_FOLDER, filename)
            print(f"   Using alternative filename: {filename}")
    
    # Test write permissions
    try:
        test_file = full_path + ".test"
        with open(test_file, 'w') as f:
            f.write("test")
        os.remove(test_file)
        print(f"   ✅ Write permissions confirmed")
    except Exception as e:
        print(f"   ❌ Write permission test failed: {e}")
        raise Exception(f"Cannot write to {SAVE_FOLDER}")
    
    # Clean up any old captured files in the directory (updated pattern)
    try:
        for file in os.listdir(SAVE_FOLDER):
            if file.startswith("Captured_") and file.endswith(".xlsx"):
                file_path = os.path.join(SAVE_FOLDER, file)
                # Remove files older than 1 hour to prevent accumulation
                if os.path.getmtime(file_path) < time.time() - 3600:
                    try:
                        os.remove(file_path)
                        print(f"   🗑️ Cleaned old file: {file}")
                    except:
                        pass
    except Exception as e:
        print(f"   ⚠️ Cleanup warning: {e}")
    
    return full_path

def save_book1_via_dde(target_title, target_hwnd):
    """Use DDE to save target workbook to target location."""
    
    # Prepare clean save location with new naming pattern
    try:
        full_path = prepare_save_location(target_title)
    except Exception as e:
        print(f"❌ Save location preparation failed: {e}")
        return None
    
    print(f"🔄 Attempting DDE save to: {full_path}")
    
    try:
        # Ensure target workbook is visible and active
        print(f"   Ensuring target window is active (HWND {target_hwnd:#010x})")
        win32gui.ShowWindow(target_hwnd, 9)  # SW_RESTORE (should already be visible)
        time.sleep(0.5)
        
        # Create DDE server
        server = dde.CreateServer()
        server.Create("TestClient")
        
        # Now try DDE - should connect to target since other Excel windows are minimized
        try:
            print(f"   Connecting to Excel|System (target should be primary)")
            conversation = dde.CreateConversation(server)
            conversation.ConnectTo("Excel", "System")
            
            print(f"   ✅ Connected to Excel DDE")
            
            # Try save commands with better error handling and format specification
            save_commands = [
                f'[SAVE.AS("{full_path}",51)]',  # 51 = Excel .xlsx format (avoids compatibility issues)
                f'[SAVE.AS("{full_path}",1)]',   # 1 = Excel format
                f'[SAVE.AS("{full_path}")]',     # Standard save
                f'[FILE.SAVE.AS("{full_path}",51)]',  # File menu with xlsx format
            ]
            
            for cmd in save_commands:
                try:
                    print(f"     Sending save command: {cmd}")
                    result = conversation.Exec(cmd)
                    
                    if result:
                        print(f"     ✅ DDE save command executed!")
                        
                        # Give Excel time to save and handle any dialogs
                        time.sleep(4)  # Longer wait for file operations
                        
                        # Check if file was created
                        if os.path.exists(full_path):
                            # Verify file is not empty and accessible
                            try:
                                file_size = os.path.getsize(full_path)
                                if file_size > 0:
                                    print(f"   ✅ File saved successfully: {full_path}")
                                    print(f"   File size: {file_size} bytes")
                                    try:
                                        conversation.Close()
                                    except AttributeError:
                                        print("     ⚠️ DDE conversation close method not available (library issue)")
                                    except Exception as e:
                                        print(f"     ⚠️ Error closing DDE conversation: {e}")
                                    try:
                                        server.Shutdown()
                                    except Exception as e:
                                        print(f"   ⚠️ Error shutting down DDE server: {e}")
                                    return full_path
                                else:
                                    print(f"     ⚠️ File created but is empty (0 bytes)")
                            except Exception as e:
                                print(f"     ⚠️ File created but cannot access: {e}")
                        else:
                            print(f"     ⚠️ Command succeeded but file not found")
                    else:
                        print(f"     ❌ Save command failed")
                        
                except Exception as e:
                    print(f"     ❌ Save command error: {e}")
                    continue
            
            try:
                conversation.Close()
            except AttributeError:
                # Handle the 'PyDDEConv' object has no attribute 'Close' error
                print("     ⚠️ DDE conversation close method not available (library issue)")
            except Exception as e:
                print(f"     ⚠️ Error closing DDE conversation: {e}")
            
        except Exception as e:
            print(f"   ❌ DDE connection failed: {e}")
        
        try:
            server.Shutdown()
        except Exception as e:
            print(f"   ⚠️ Error shutting down DDE server: {e}")
            
        print("   ❌ All DDE save attempts failed")
        return None
        
    except Exception as e:
        print(f"   ❌ DDE setup failed: {e}")
        return None

def simple_keyboard_save_approach(target_hwnd):
    """Fallback: Use keyboard automation to save target workbook."""
    
    print(f"🔄 Trying keyboard automation approach...")
    
    try:
        # Activate target window
        print(f"   Activating target window...")
        win32gui.ShowWindow(target_hwnd, 9)  # SW_RESTORE  
        time.sleep(1)
        
        # Send Ctrl+S to save
        print(f"   Sending Ctrl+S...")
        
        # Send Ctrl down
        win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        # Send S down  
        win32gui.PostMessage(target_hwnd, win32con.WM_KEYDOWN, ord('S'), 0)
        # Send S up
        win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, ord('S'), 0)
        # Send Ctrl up
        win32gui.PostMessage(target_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        
        time.sleep(2)  # Wait for save dialog
        
        print(f"   Save dialog should be open - target workbook is now being saved")
        print(f"   (Manual save required, or dialog can be automated further)")
        
        return "SAVE_TRIGGERED"
        
    except Exception as e:
        print(f"   ❌ Keyboard save failed: {e}")
        return None

def main():
    print("Excel DDE Auto-Saver with Window Minimization")
    print("=" * 50)
    
    # Show all Excel windows
    print("Current Excel windows:")
    excel_windows = list_excel_windows()
    for pid, hwnd, title, vis in excel_windows:
        if vis:
            exe = psutil.Process(pid).name()
            if exe.lower() == "excel.exe":
                print(f"  PID {pid:>5}: {title}")
    
    print("\n" + "─" * 50)
    
    # Find unsaved workbook (Book1, Book2, etc.) - improved detection
    target_pid, target_hwnd, target_title = find_book1_window()
    
    if not target_pid:
        print("❌ No unsaved workbook found")
        print("   Looking for windows like 'Book1 - Excel', 'Book2 - Excel', etc.")
        print("   (Avoiding our own 'Captured_' files)")
        return
    
    workbook_name = extract_workbook_name(target_title)
    print(f"🎯 Found unsaved workbook:")
    print(f"   PID: {target_pid}")
    print(f"   HWND: {target_hwnd:#010x}")
    print(f"   Title: {target_title}")
    print(f"   Workbook: {workbook_name}")
    
    print("\n" + "─" * 50)
    
    # CRITICAL: Minimize other Excel windows to make target primary
    minimized_windows = minimize_other_excel_windows(target_hwnd)
    
    try:
        # Try DDE save with target as the only visible Excel
        print(f"\nAttempting to save {workbook_name} via DDE...")
        saved_file = save_book1_via_dde(target_title, target_hwnd)
        
        # Method 2: Keyboard automation if DDE failed
        if not saved_file:
            saved_file = simple_keyboard_save_approach(target_hwnd)
        
        # Report results
        if saved_file and saved_file != "SAVE_TRIGGERED":
            print(f"\n✅ SUCCESS: {workbook_name} saved to {saved_file}")
            
            # Enhanced verification
            try:
                file_size = os.path.getsize(saved_file)
                print(f"File size: {file_size} bytes")
                
                if file_size > 1000:  # Reasonable minimum for Excel file
                    # Verify it's a valid Excel file
                    import pandas as pd
                    df = pd.read_excel(saved_file)
                    print(f"Content verification: {len(df)} rows, {len(df.columns)} columns")
                    
                    # Show sample of first few values to confirm it's the right data
                    if not df.empty:
                        print(f"Sample data preview:")
                        for col in list(df.columns)[:3]:  # Show first 3 columns
                            sample_vals = df[col].dropna().head(2).tolist()
                            print(f"  {col}: {sample_vals}")
                else:
                    print(f"⚠️ File size seems too small ({file_size} bytes) - may be corrupted")
                    
            except Exception as e:
                print(f"⚠️ File saved but verification failed: {e}")
                print("File may still be valid - check manually")
                
        elif saved_file == "SAVE_TRIGGERED":
            print(f"\n⚠️ Save process triggered but file path unknown")
            print(f"{workbook_name} save dialog was opened - manual completion may be required")
        else:
            print(f"\n❌ Could not save {workbook_name} via DDE or keyboard automation")
            print("Target Excel instance may not be responding to automation")
    
    finally:
        # Always restore minimized windows, regardless of success/failure
        print("\n" + "─" * 50)
        restore_minimized_windows(minimized_windows)
        print("✅ Cleanup completed")

if __name__ == "__main__":
    main()