"""
autosaver.py - Reliable Excel Book1 Capture Module
Based on proven DDE + foreground approach
"""

import time
import win32ui
import win32gui
import win32process
import psutil
import dde
import os
from datetime import datetime

def list_excel_windows():
    """Return [(pid, hwnd, title, visible)] for every XLMAIN window."""
    windows = []
    def _enum(hwnd, _):
        if win32gui.GetClassName(hwnd) == "XLMAIN":
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            title = win32gui.GetWindowText(hwnd)
            vis = win32gui.IsWindowVisible(hwnd)
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

def bring_to_foreground(target_hwnd, verbose=True):
    """Bring target window to foreground."""
    if verbose:
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
            if verbose:
                print(f"   ✅ Book1 brought to foreground")
            return True
        else:
            if verbose:
                print(f"   ⚠️ Foreground operation may not have worked")
            return False
            
    except Exception as e:
        if verbose:
            print(f"   ❌ Error bringing to foreground: {e}")
        return False

def save_book1_dde(target_title, save_folder, filename=None, verbose=True):
    """Save Book1 using DDE to specified location."""
    
    # Ensure save folder exists
    os.makedirs(save_folder, exist_ok=True)
    
    # Generate filename if not provided
    if filename is None:
        timestamp = datetime.now().strftime("%m-%d-%Y_%H.%M")  # Match your original format
        filename = f"Captured_{timestamp}.xlsx"
    
    full_path = os.path.join(save_folder, filename)
    
    if verbose:
        print(f"💾 Saving Book1 to: {full_path}")
    
    # Clean up existing file
    if os.path.exists(full_path):
        try:
            os.remove(full_path)
            if verbose:
                print(f"   🗑️ Removed existing file")
        except:
            pass
    
    try:
        # Create DDE server
        server = dde.CreateServer()
        server.Create("CaptureClient")
        
        # Connect to Excel
        conversation = dde.CreateConversation(server)
        conversation.ConnectTo("Excel", "System")
        
        if verbose:
            print(f"   ✅ DDE Connected")
        
        # Simple save operation
        result = conversation.Exec(f'[SAVE.AS("{full_path}")]')
        if verbose:
            print(f"   📤 Save command sent")
        
        # Wait and check result
        time.sleep(2)
        if os.path.exists(full_path):
            file_size = os.path.getsize(full_path)
            if file_size > 0:
                if verbose:
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
                if verbose:
                    print(f"   ❌ File created but empty")
        else:
            if verbose:
                print(f"   ❌ File not created")
        
        # Close DDE
        try:
            conversation.Close()
        except:
            pass
        server.Shutdown()
        return None
        
    except Exception as e:
        if verbose:
            print(f"   ❌ DDE save failed: {e}")
        return None

def capture_book1(save_folder, filename=None, verbose=True):
    """
    Main function: Capture Book1 workbook to specified folder.
    
    Args:
        save_folder (str): Directory to save the captured file
        filename (str, optional): Custom filename. If None, auto-generates with timestamp
        verbose (bool): Whether to print status messages
    
    Returns:
        str: Path to saved file on success, None on failure
    """
    
    if verbose:
        print("🔍 Looking for Book1 workbook...")
    
    # Find target Book1
    target_result = find_book1_window_filtered()
    if not target_result[0]:
        if verbose:
            print("❌ No Book1 found (excluding captured files)")
        return None
    
    target_pid, target_hwnd, target_title = target_result
    
    if verbose:
        print(f"🎯 Found Book1: {target_title}")
    
    # Bring to foreground
    success = bring_to_foreground(target_hwnd, verbose)
    if not success and verbose:
        print("⚠️ Warning: Foreground operation failed, abortion...")
        return
    
    # Save using DDE
    saved_file = save_book1_dde(target_title, save_folder, filename, verbose)
    
    if saved_file and verbose:
        print(f"✅ Book1 captured successfully!")
        print(f"📁 Saved to: {saved_file}")
    elif not saved_file and verbose:
        print("❌ Failed to capture Book1")
    
    return saved_file

def is_book1_available():
    """
    Check if Book1 workbook is available for capture.
    
    Returns:
        bool: True if Book1 is available, False otherwise
    """
    target_result = find_book1_window_filtered()
    return target_result[0] is not None

def get_available_workbooks():
    """
    Get list of available unsaved workbook names.
    
    Returns:
        list: List of workbook titles available for capture
    """
    excel_windows = list_excel_windows()
    workbooks = []
    
    for pid, hwnd, title, vis in excel_windows:
        if vis:
            try:
                exe = psutil.Process(pid).name()
                if exe.lower() == "excel.exe":
                    # Extract workbook name from title
                    if " - Excel" in title:
                        workbook_name = title.replace(" - Excel", "").strip()
                        # Only include unsaved workbooks (Book1, Book2, etc.)
                        if workbook_name.lower().startswith('book') and not workbook_name.lower().startswith('captured_'):
                            workbooks.append(workbook_name)
            except:
                continue
    
    return workbooks

# Standalone test function
def test_capture():
    """Test the capture functionality."""
    print("Excel Book1 Auto-Saver Test")
    print("=" * 30)
    
    # Show current Excel windows
    print("Current Excel windows:")
    excel_windows = list_excel_windows()
    for pid, hwnd, title, vis in excel_windows:
        if vis:
            try:
                exe = psutil.Process(pid).name()
                if exe.lower() == "excel.exe":
                    print(f"  PID {pid:>5}: {title}")
            except:
                continue
    
    print("\n" + "─" * 30)
    
    # Test capture
    save_folder = r"C:\Users\sasuk\Documents\CapturedExports"
    saved_file = capture_book1(save_folder, verbose=True)
    
    if saved_file:
        print(f"\n✅ SUCCESS: Book1 captured!")
        print(f"📁 Saved to: {saved_file}")
    else:
        print(f"\n❌ Failed to save Book1")
    
    print("\n✅ Test completed")

if __name__ == "__main__":
    # Run test when executed directly
    test_capture()