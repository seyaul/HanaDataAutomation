import time, win32gui, win32process, psutil
import win32ui
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

def find_book1_window_filtered():
    """Find Book1 Excel window, excluding captured files."""
    excel_windows = list_excel_windows()
    for pid, hwnd, title, vis in excel_windows:
        if vis and 'book1' in title.lower():
            if 'captured_' not in title.lower():
                return pid, hwnd, title
    return None, None, None

def bring_to_foreground(target_hwnd, target_title):
    """Bring target window to foreground."""
    print(f"🎯 Bringing Book1 to foreground...")
    
    try:
        win32gui.ShowWindow(target_hwnd, 9)  # SW_RESTORE
        time.sleep(0.5)
        win32gui.SetForegroundWindow(target_hwnd)
        time.sleep(0.5)
        win32gui.BringWindowToTop(target_hwnd)
        time.sleep(0.5)
        
        # Verify it worked
        current_fg = win32gui.GetForegroundWindow()
        if current_fg == target_hwnd:
            print(f"   ✅ Book1 brought to foreground")
        else:
            print(f"   ⚠️ Foreground operation may not have worked")
            
    except Exception as e:
        print(f"   ❌ Error bringing to foreground: {e}")

def save_book1_simple(target_title):
    """Simple DDE save operation using foreground-only approach."""
    
    # Ensure save folder exists
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Book1_Captured_{timestamp}.xlsx"
    full_path = os.path.join(SAVE_FOLDER, filename)
    
    print(f"💾 Saving Book1 to: {full_path}")
    
    # Clean up existing file
    if os.path.exists(full_path):
        try:
            os.remove(full_path)
            print(f"   🗑️ Removed existing file")
        except:
            pass
    
    try:
        # Create DDE server
        server = dde.CreateServer()
        server.Create("TestClient")
        
        # Connect to Excel
        conversation = dde.CreateConversation(server)
        conversation.ConnectTo("Excel", "System")
        
        print(f"   ✅ DDE Connected")
        
        # Simple save operation
        result = conversation.Exec(f'[SAVE.AS("{full_path}")]')
        print(f"   📤 Save command sent")
        
        # Wait and check result
        time.sleep(2)
        if os.path.exists(full_path):
            file_size = os.path.getsize(full_path)
            if file_size > 0:
                print(f"   ✅ File saved successfully!")
                print(f"   📊 File size: {file_size} bytes")
                
                # Verify it's a valid Excel file
                try:
                    import pandas as pd
                    df = pd.read_excel(full_path)
                    print(f"   📋 Content: {len(df)} rows, {len(df.columns)} columns")
                except Exception as e:
                    print(f"   ⚠️ Verification failed: {e}")
                
                # Close DDE
                try:
                    conversation.Close()
                except:
                    pass
                server.Shutdown()
                return full_path
            else:
                print(f"   ❌ File created but empty")
        else:
            print(f"   ❌ File not created")
        
        # Close DDE
        try:
            conversation.Close()
        except:
            pass
        server.Shutdown()
        return None
        
    except Exception as e:
        print(f"   ❌ DDE save failed: {e}")
        return None

def main():
    print("Excel Book1 Auto-Saver (Foreground-Only)")
    print("=" * 45)
    
    # Show current Excel windows
    print("Current Excel windows:")
    excel_windows = list_excel_windows()
    for pid, hwnd, title, vis in excel_windows:
        if vis:
            exe = psutil.Process(pid).name()
            if exe.lower() == "excel.exe":
                print(f"  PID {pid:>5}: {title}")
    
    # Find target Book1
    target_result = find_book1_window_filtered()
    if not target_result[0]:
        print("\n❌ No Book1 found (excluding captured files)")
        return
    
    target_pid, target_hwnd, target_title = target_result
    print(f"\n🎯 Found Book1:")
    print(f"   PID: {target_pid}")
    print(f"   Title: {target_title}")
    
    print("\n" + "─" * 45)
    
    # Execute the streamlined workflow
    bring_to_foreground(target_hwnd, target_title)
    saved_file = save_book1_simple(target_title)
    
    if saved_file:
        print(f"\n✅ SUCCESS: Book1 captured!")
        print(f"📁 Saved to: {saved_file}")
    else:
        print(f"\n❌ Failed to save Book1")
    
    print("\n✅ Process completed")

if __name__ == "__main__":
    main()