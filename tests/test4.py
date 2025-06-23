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
        # CRITICAL: First activate the Book1 window to make it the active Excel
        print(f"   Activating Book1 window (HWND {book1_hwnd:#010x})")
        win32gui.ShowWindow(book1_hwnd, 9)  # SW_RESTORE
        win32gui.SetActiveWindow(book1_hwnd)
        time.sleep(1)  # Give time for activation
        
        # Create DDE server
        server = dde.CreateServer()
        server.Create("TestClient")
        
        # Now try DDE with Book1 as the active window
        try:
            print(f"   Connecting to Excel|System (Book1 should now be active)")
            conversation = dde.CreateConversation(server)
            conversation.ConnectTo("Excel", "System")
            
            print(f"   ✅ Connected to Excel DDE")
            
            # Try Book1-specific commands first
            book1_commands = [
                f'[ACTIVATE("Book1")]',  # Ensure Book1 is active
                f'[SELECT("Book1")]',    # Select Book1 workbook
                f'[WORKBOOK.ACTIVATE("Book1")]',  # Alternative activation
            ]
            
            # Try to activate Book1 first
            for cmd in book1_commands:
                try:
                    print(f"     Activating Book1: {cmd}")
                    result = conversation.Exec(cmd)
                    if result:
                        print(f"     ✅ Book1 activation command succeeded")
                        break
                except Exception as e:
                    print(f"     ⚠️ Activation command failed: {e}")
                    continue
            
            # Now try save commands with Book1-specific syntax
            save_commands = [
                f'[SAVE.AS("Book1","{full_path}")]',  # Specify workbook name
                f'[SAVE.AS("Book1","{full_path}",1)]',  # With file type
                f'[SAVE.AS("{full_path}")]',  # Standard save
                f'[FILE.SAVE.AS("{full_path}",1)]',  # File menu save
                f'[WORKBOOK.SAVE.AS("Book1","{full_path}")]',  # Workbook-specific
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
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Book1_Keyboard_{timestamp}.xlsx"
        full_path = os.path.join(SAVE_FOLDER, filename)
        
        # Activate Book1 window
        print(f"   Activating Book1 window...")
        win32gui.ShowWindow(book1_hwnd, 9)  # SW_RESTORE  
        win32gui.SetActiveWindow(book1_hwnd)
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
        
        # This will open Save As dialog for unsaved workbooks
        # At this point, Book1 should show a save dialog
        print(f"   Save dialog should be open - Book1 is now being saved")
        print(f"   (Manual save required, or dialog can be automated further)")
        
        # For now, just return that we triggered the save process
        return "SAVE_TRIGGERED"
        
    except Exception as e:
        print(f"   ❌ Keyboard save failed: {e}")
        return None

def main():
    print("Excel DDE Auto-Saver")
    print("=" * 40)
    
    # Show all Excel windows
    print("Current Excel windows:")
    excel_windows = list_excel_windows()
    for pid, hwnd, title, vis in excel_windows:
        if vis:
            exe = psutil.Process(pid).name()
            if exe.lower() == "excel.exe":
                print(f"  PID {pid:>5}: {title}")
    
    print("\n" + "─" * 40)
    
    # Find Book1 specifically
    book1_pid, book1_hwnd, book1_title = find_book1_window()
    
    if not book1_pid:
        print("❌ No Book1 window found")
        return
    
    print(f"🎯 Found Book1:")
    print(f"   PID: {book1_pid}")
    print(f"   HWND: {book1_hwnd:#010x}")
    print(f"   Title: {book1_title}")
    
    print("\n" + "─" * 40)
    
    # Try DDE save
    print("Attempting to save Book1 via DDE...")
    
    # Method 1: Standard DDE (with window activation)
    saved_file = save_book1_via_dde(book1_title, book1_hwnd)
    
    # Method 2: Keyboard automation if DDE failed
    # if not saved_file:
    #     saved_file = simple_keyboard_save_approach(book1_hwnd)
    #     pass
    
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

if __name__ == "__main__":
    main()