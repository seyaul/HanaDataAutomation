import win32com.client
import win32gui
import win32process
import psutil
import pythoncom
import time

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

def get_excel_app_by_pid(target_pid):
    """Get Excel application object for a specific PID using ROT."""
    try:
        pythoncom.CoInitialize()
        
        # Get the Running Object Table
        rot = pythoncom.GetRunningObjectTable()
        enum_moniker = rot.EnumRunning()
        
        for moniker in enum_moniker:
            try:
                # Get the display name of the object
                display_name = moniker.GetDisplayName(None, None)
                
                # Look for Excel application objects
                if 'Excel.Application' in display_name:
                    # Get the actual object
                    obj = rot.GetObject(moniker)
                    app = win32com.client.Dispatch(obj)
                    
                    # Check if this app belongs to our target PID
                    app_hwnd = app.Hwnd
                    _, app_pid = win32process.GetWindowThreadProcessId(app_hwnd)
                    
                    if app_pid == target_pid:
                        print(f"✅ Found Excel app for PID {target_pid}")
                        return app
                        
            except Exception as e:
                continue
                
        print(f"❌ Could not find Excel app for PID {target_pid}")
        return None
        
    except Exception as e:
        print(f"⚠️ Error accessing ROT: {e}")
        return None

def find_book1_workbook():
    """Find Book1 by targeting the specific Excel instance."""
    
    # First, find the Excel window with "Book1" in the title
    excel_windows = list_excel_windows()
    book1_pid = None
    book1_hwnd = None
    
    for pid, hwnd, title, vis in excel_windows:
        if vis and 'book1' in title.lower():
            book1_pid = pid
            book1_hwnd = hwnd
            print(f"🎯 Found Book1 window: PID {pid}, HWND {hwnd:#010x}, Title: '{title}'")
            break
    
    if not book1_pid:
        print("❌ No Book1 window found")
        return None
    
    # Now get the Excel application for that specific PID
    excel_app = get_excel_app_by_pid(book1_pid)
    
    if not excel_app:
        print(f"❌ Could not connect to Excel instance with PID {book1_pid}")
        return None
    
    # Find the Book1 workbook in this specific Excel instance
    try:
        print(f"📋 Excel instance has {excel_app.Workbooks.Count} workbook(s)")
        
        for i in range(1, excel_app.Workbooks.Count + 1):
            workbook = excel_app.Workbooks.Item(i)
            workbook_name = workbook.Name
            print(f"  Workbook {i}: '{workbook_name}' (Path: '{workbook.Path}')")
            
            # Check if this is Book1 (unsaved workbooks typically have empty path)
            if (workbook_name.lower().startswith('book') and 
                workbook.Path == ""):  # Empty path indicates unsaved workbook
                
                print(f"🎯 SUCCESS: Found Book1 workbook!")
                print(f"   Name: {workbook_name}")
                print(f"   Full path: {workbook.FullName}")
                print(f"   Saved: {workbook.Saved}")
                print(f"   Sheets: {workbook.Sheets.Count}")
                
                return workbook
        
        print("❌ No Book1-like workbook found in the target Excel instance")
        return None
        
    except Exception as e:
        print(f"⚠️ Error accessing workbooks: {e}")
        return None

def verify_workbook(workbook):
    """Simple tests to verify you have the right workbook."""
    try:
        print("\n🔍 Workbook Verification:")
        print(f"   Name: {workbook.Name}")
        print(f"   Path: '{workbook.Path}' (empty = unsaved)")
        print(f"   Saved status: {workbook.Saved}")
        print(f"   Number of sheets: {workbook.Sheets.Count}")
        
        # Check first sheet
        if workbook.Sheets.Count > 0:
            first_sheet = workbook.Sheets(1)
            print(f"   First sheet name: '{first_sheet.Name}'")
            
            # Check if it has data
            used_range = first_sheet.UsedRange
            if used_range:
                print(f"   Data range: {used_range.Address}")
                print(f"   Rows with data: {used_range.Rows.Count}")
                print(f"   Columns with data: {used_range.Columns.Count}")
            else:
                print("   No data found in first sheet")
        
        # Get the Excel application info
        app = workbook.Application
        print(f"   Excel PID: {win32process.GetWindowThreadProcessId(app.Hwnd)[1]}")
        
        return True
        
    except Exception as e:
        print(f"⚠️ Error during verification: {e}")
        return False

if __name__ == "__main__":
    print("Searching for Book1 workbook...")
    print("─" * 50)
    
    # Show all Excel windows first
    print("Excel windows found:")
    for pid, hwnd, title, vis in list_excel_windows():
        if vis:
            print(f"  PID {pid}: '{title}'")
    
    print("\n" + "─" * 50)
    
    # Find Book1
    book1_workbook = find_book1_workbook()
    
    if book1_workbook:
        verify_workbook(book1_workbook)
        print(f"\n✅ Book1 workbook ready for processing!")
        
        # Your existing processing logic can go here
        # Example: book1_workbook.SaveAs(r"C:\path\to\save\Book1.xlsx")
        
    else:
        print("❌ Could not find or access Book1 workbook")