import time, win32gui, win32process, psutil
import win32ui, win32con
import dde
import os
from datetime import datetime

# Target directory
SAVE_FOLDER = r"C:\Users\sasuk\Documents\CapturedExports"

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
    """Find the Book1 Excel window."""
    excel_windows = list_excel_windows()
    
    for pid, hwnd, title, vis in excel_windows:
        if vis and 'book1' in title.lower():
            return pid, hwnd, title
    return None, None, None

def minimize_other_excel_windows(book1_hwnd):
    """Minimize all Excel windows except Book1 to make it primary."""
    
    print("🔽 Minimizing other Excel windows to make Book1 primary...")
    
    excel_windows = list_excel_windows()
    minimized_windows = []
    
    for pid, hwnd, title, vis in excel_windows:
        if vis and hwnd != book1_hwnd:  # Skip Book1 itself
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

def save_book1_via_dde(book1_title, book1_hwnd):
    """Use DDE to save Book1 to target location."""
    
    # Ensure save folder exists
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Book1_Captured_{timestamp}.xlsx"
    full_path = os.path.join(SAVE_FOLDER, filename)
    
    print(f"🔄 Attempting DDE save to: {full_path}")
    
    try:
        # Ensure Book1 is visible and active
        print(f"   Ensuring Book1 window is active (HWND {book1_hwnd:#010x})")
        win32gui.ShowWindow(book1_hwnd, 9)  # SW_RESTORE (should already be visible)
        time.sleep(0.5)
        
        # Create DDE server
        server = dde.CreateServer()
        server.Create("TestClient")
        
        # Now try DDE - should connect to Book1 since other Excel windows are minimized
        try:
            print(f"   Connecting to Excel|System (Book1 should be primary)")
            conversation = dde.CreateConversation(server)
            conversation.ConnectTo("Excel", "System")
            
            print(f"   ✅ Connected to Excel DDE")
            
            # Try save commands - simplified since Book1 should be the only active Excel
            save_commands = [
                f'[SAVE.AS("{full_path}")]',  # Standard save
                f'[SAVE.AS("{full_path}",1)]',  # With file type
                f'[FILE.SAVE.AS("{full_path}",1)]',  # File menu save
            ]
            
            for cmd in save_commands:
                try:
                    print(f"     Sending save command: {cmd}")
                    result = conversation.Exec(cmd)
                    
                    if result:
                        print(f"     ✅ DDE save command executed!")
                        
                        # Check if file was created
                        time.sleep(3)  # Give Excel more time to save
                        if os.path.exists(full_path):
                            print(f"   ✅ File saved successfully: {full_path}")
                            conversation.Close()
                            server.Shutdown()
                            return full_path
                        else:
                            print(f"     ⚠️ Command succeeded but file not found")
                    else:
                        print(f"     ❌ Save command failed")
                        
                except Exception as e:
                    print(f"     ❌ Save command error: {e}")
                    continue
            
            conversation.Close()
            
        except Exception as e:
            print(f"   ❌ DDE connection failed: {e}")
        
        server.Shutdown()
        print("   ❌ All DDE save attempts failed")
        return None
        
    except Exception as e:
        print(f"   ❌ DDE setup failed: {e}")
        return None

def simple_keyboard_save_approach(book1_hwnd):
    """Fallback: Use keyboard automation to save Book1."""
    
    print(f"🔄 Trying keyboard automation approach...")
    
    try:
        # Activate Book1 window
        print(f"   Activating Book1 window...")
        win32gui.ShowWindow(book1_hwnd, 9)  # SW_RESTORE  
        time.sleep(1)
        
        # Send Ctrl+S to save
        print(f"   Sending Ctrl+S...")
        
        # Send Ctrl down
        win32gui.PostMessage(book1_hwnd, win32con.WM_KEYDOWN, win32con.VK_CONTROL, 0)
        # Send S down  
        win32gui.PostMessage(book1_hwnd, win32con.WM_KEYDOWN, ord('S'), 0)
        # Send S up
        win32gui.PostMessage(book1_hwnd, win32con.WM_KEYUP, ord('S'), 0)
        # Send Ctrl up
        win32gui.PostMessage(book1_hwnd, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
        
        time.sleep(2)  # Wait for save dialog
        
        print(f"   Save dialog should be open - Book1 is now being saved")
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
    
    # Find Book1 specifically
    book1_pid, book1_hwnd, book1_title = find_book1_window()
    
    if not book1_pid:
        print("❌ No Book1 window found")
        return
    
    print(f"🎯 Found Book1:")
    print(f"   PID: {book1_pid}")
    print(f"   HWND: {book1_hwnd:#010x}")
    print(f"   Title: {book1_title}")
    
    print("\n" + "─" * 50)
    
    # CRITICAL: Minimize other Excel windows to make Book1 primary
    minimized_windows = minimize_other_excel_windows(book1_hwnd)
    
    try:
        # Try DDE save with Book1 as the only visible Excel
        print("\nAttempting to save Book1 via DDE...")
        saved_file = save_book1_via_dde(book1_title, book1_hwnd)
        
        # Method 2: Keyboard automation if DDE failed
        # if not saved_file:
        #     saved_file = simple_keyboard_save_approach(book1_hwnd)
        
        # Report results
        if saved_file and saved_file != "SAVE_TRIGGERED":
            print(f"\n✅ SUCCESS: Book1 saved to {saved_file}")
            print(f"File size: {os.path.getsize(saved_file)} bytes")
            
            # Verify it's a valid Excel file
            try:
                import pandas as pd
                df = pd.read_excel(saved_file)
                print(f"Verification: Excel file has {len(df)} rows, {len(df.columns)} columns")
            except Exception as e:
                print(f"⚠️ File saved but verification failed: {e}")
        elif saved_file == "SAVE_TRIGGERED":
            print(f"\n⚠️ Save process triggered but file path unknown")
            print("Book1 save dialog was opened - manual completion may be required")
        else:
            print("\n❌ Could not save Book1 via DDE or keyboard automation")
            print("Book1 Excel instance may not be responding to automation")
    
    finally:
        # Always restore minimized windows, regardless of success/failure
        print("\n" + "─" * 50)
        restore_minimized_windows(minimized_windows)
        print("✅ Cleanup completed")

if __name__ == "__main__":
    main()