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

def find_book1_window():
    """Find the Book1 Excel window."""
    excel_windows = list_excel_windows()
    
    for pid, hwnd, title, vis in excel_windows:
        if vis and 'book1' in title.lower():
            return pid, hwnd, title
    return None, None, None

def save_book1_via_dde(book1_title):
    """Use DDE to save Book1 to target location."""
    
    # Ensure save folder exists
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Book1_Captured_{timestamp}.xlsx"
    full_path = os.path.join(SAVE_FOLDER, filename)
    
    print(f"🔄 Attempting DDE save to: {full_path}")
    
    try:
        # Create DDE server
        server = dde.CreateServer()
        server.Create("TestClient")
        
        # Try different DDE connection approaches
        dde_topics = ["System", "[Book1]Book1", "Book1", "Sheet1"]
        
        for topic in dde_topics:
            try:
                print(f"   Trying DDE topic: 'Excel|{topic}'")
                
                # Connect to Excel with this topic
                conversation = dde.CreateConversation(server)
                conversation.ConnectTo("Excel", topic)
                
                print(f"   ✅ Connected to Excel via DDE topic: {topic}")
                
                # Try to save using DDE EXECUTE command
                # Excel DDE commands use square bracket syntax
                save_commands = [
                    f'[SAVE.AS("{full_path}")]',
                    f'[SAVE.AS("{full_path}",1)]',  # 1 = Excel format
                    f'[FILE.SAVE.AS("{full_path}")]',
                ]
                
                for cmd in save_commands:
                    try:
                        print(f"     Sending DDE command: {cmd}")
                        result = conversation.Exec(cmd)
                        
                        if result:
                            print(f"   ✅ DDE command executed successfully!")
                            
                            # Check if file was created
                            time.sleep(2)  # Give Excel time to save
                            if os.path.exists(full_path):
                                print(f"   ✅ File saved successfully: {full_path}")
                                conversation.Close()
                                server.Shutdown()
                                return full_path
                            else:
                                print(f"   ⚠️ DDE command succeeded but file not found")
                        else:
                            print(f"     ❌ DDE command failed")
                            
                    except Exception as e:
                        print(f"     ❌ DDE command error: {e}")
                        continue
                
                conversation.Close()
                
            except Exception as e:
                print(f"   ❌ Could not connect to topic '{topic}': {e}")
                continue
        
        server.Shutdown()
        print("   ❌ All DDE connection attempts failed")
        return None
        
    except Exception as e:
        print(f"   ❌ DDE setup failed: {e}")
        return None

def alternative_dde_approach(book1_hwnd, book1_pid):
    """Alternative DDE approach using window-specific connection."""
    
    print(f"🔄 Trying alternative DDE approach for HWND {book1_hwnd:#010x}")
    
    try:
        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Book1_Alt_{timestamp}.xlsx"
        full_path = os.path.join(SAVE_FOLDER, filename)
        
        # Try using win32ui DDE functions
        import win32ui
        
        # Create DDE client
        dde_client = win32ui.CreateDDEClient()
        
        # Try to connect to Excel
        try:
            # Different service names Excel might respond to
            service_names = ["Excel", "EXCEL", f"Excel.{book1_pid}"]
            
            for service in service_names:
                try:
                    print(f"   Trying DDE service: {service}")
                    dde_server = dde_client.CreateConversation(service)
                    
                    # Try different topic names
                    topics = ["System", "Book1", "[Book1]", f"[Book1]Sheet1"]
                    
                    for topic in topics:
                        try:
                            print(f"     Connecting to topic: {topic}")
                            dde_server.ConnectTo(topic)
                            
                            print(f"   ✅ Connected via {service}|{topic}")
                            
                            # Send save command
                            save_cmd = f'[SAVE.AS("{full_path}",1)]'
                            print(f"     Executing: {save_cmd}")
                            
                            result = dde_server.Exec(save_cmd)
                            
                            if result:
                                time.sleep(2)
                                if os.path.exists(full_path):
                                    print(f"   ✅ Alternative DDE save successful!")
                                    dde_server.Close()
                                    return full_path
                            
                            dde_server.Close()
                            
                        except Exception as e:
                            print(f"       Topic {topic} failed: {e}")
                            continue
                            
                except Exception as e:
                    print(f"     Service {service} failed: {e}")
                    continue
        
        except Exception as e:
            print(f"   ❌ Alternative DDE failed: {e}")
            
        return None
        
    except Exception as e:
        print(f"   ❌ Alternative DDE setup failed: {e}")
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
    
    # Method 1: Standard DDE
    saved_file = save_book1_via_dde(book1_title)
    
    # Method 2: Alternative DDE if first failed
    if not saved_file:
        saved_file = alternative_dde_approach(book1_hwnd, book1_pid)
    
    if saved_file:
        print(f"\n✅ SUCCESS: Book1 saved to {saved_file}")
        print(f"File size: {os.path.getsize(saved_file)} bytes")
        
        # Verify it's a valid Excel file
        try:
            import pandas as pd
            df = pd.read_excel(saved_file)
            print(f"Verification: Excel file has {len(df)} rows, {len(df.columns)} columns")
        except Exception as e:
            print(f"⚠️ File saved but verification failed: {e}")
    else:
        print("\n❌ Could not save Book1 via DDE")
        print("Book1 Excel instance may not support DDE or is not responding")

if __name__ == "__main__":
    main()