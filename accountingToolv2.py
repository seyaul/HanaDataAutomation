import win32com.client
import win32gui
import win32process
import win32con
import psutil
import os
import time
import pandas as pd
import subprocess
from typing import List, Dict, Tuple

SAVE_FOLDER = r"C:\Users\sasuk\Documents\CapturedExports"
PROCESSED_FOLDER = r"C:\Users\sasuk\Documents\ProcessedExports"
BRAND_MAP_CSV = "brand_map.csv"

os.makedirs(SAVE_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

processed_workbooks = set()

# ──────────────────────────────────────────────────────────────────────────

def find_unsaved_excel_windows():
    """Find all Excel windows with unsaved workbooks"""
    excel_windows = []
    
    def enum_callback(hwnd, _):
        if (win32gui.GetClassName(hwnd) == "XLMAIN" and 
            win32gui.IsWindowVisible(hwnd)):
            
            window_title = win32gui.GetWindowText(hwnd)
            if window_title.lower().startswith("book"):
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    excel_windows.append((hwnd, window_title, pid))
                except:
                    pass
    
    win32gui.EnumWindows(enum_callback, None)
    return excel_windows

def get_all_excel_processes():
    """Get all Excel process PIDs"""
    excel_pids = []
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] == 'EXCEL.EXE':
            excel_pids.append(proc.info['pid'])
    return excel_pids

def get_all_excel_window_info(target_pid: int) -> List[Dict]:
    """
    Get information about all Excel windows and their files
    Returns list of dicts with window info and file paths
    """
    excel_windows = []
    
    def enum_callback(hwnd, _):
        if (win32gui.GetClassName(hwnd) == "XLMAIN" and 
            win32gui.IsWindowVisible(hwnd)):
            
            try:
                window_title = win32gui.GetWindowText(hwnd)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                
                # Skip the target process
                if pid == target_pid:
                    return
                
                excel_windows.append({
                    'hwnd': hwnd,
                    'pid': pid,
                    'title': window_title,
                    'files': []
                })
            except:
                pass
    
    win32gui.EnumWindows(enum_callback, None)
    return excel_windows

def reliable_window_activation(hwnd: int, max_attempts: int = 3) -> bool:
    """
    Try multiple techniques to activate an Excel window reliably
    """
    for attempt in range(max_attempts):
        try:
            # Method A: Basic activation
            if attempt == 0:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.1)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)
            
            # Method B: Force to top then activate
            elif attempt == 1:
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0,
                                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW)
                time.sleep(0.1)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.1)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)
            
            # Method C: Minimize others, then restore target
            else:
                # Minimize all other Excel windows first
                def minimize_others(hwnd_inner, _):
                    if (hwnd_inner != hwnd and 
                        win32gui.GetClassName(hwnd_inner) == "XLMAIN" and 
                        win32gui.IsWindowVisible(hwnd_inner)):
                        try:
                            win32gui.ShowWindow(hwnd_inner, win32con.SW_MINIMIZE)
                        except:
                            pass
                
                win32gui.EnumWindows(minimize_others, None)
                time.sleep(0.2)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.1)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)
            
            # Test if activation worked
            try:
                current_foreground = win32gui.GetForegroundWindow()
                if current_foreground == hwnd:
                    print(f"         ✅ Window activated successfully (method {attempt + 1})")
                    return True
                else:
                    print(f"         ⚠️  Activation attempt {attempt + 1} failed")
            except:
                print(f"         ⚠️  Could not verify activation {attempt + 1}")
            
        except Exception as e:
            print(f"         ⚠️  Activation method {attempt + 1} error: {e}")
    
    print(f"         ❌ All activation methods failed for window {hwnd}")
    return False

def get_all_excel_window_info(target_pid: int) -> List[Dict]:
    """
    Get information about ALL Excel windows, including protected view and multiple PIDs
    """
    excel_windows = []
    all_pids = set()
    
    def enum_callback(hwnd, _):
        if (win32gui.GetClassName(hwnd) == "XLMAIN" and 
            win32gui.IsWindowVisible(hwnd)):
            
            try:
                window_title = win32gui.GetWindowText(hwnd)
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                all_pids.add(pid)
                
                # Skip the target process (the one with Book1)
                if pid == target_pid:
                    return
                
                # Detect protected view and other special cases
                is_protected_view = "[Protected View]" in window_title
                is_read_only = "[Read-Only]" in window_title
                is_compatibility_mode = "[Compatibility Mode]" in window_title
                
                excel_windows.append({
                    'hwnd': hwnd,
                    'pid': pid,
                    'title': window_title,
                    'files': [],
                    'is_protected_view': is_protected_view,
                    'is_read_only': is_read_only,
                    'is_compatibility_mode': is_compatibility_mode
                })
            except:
                pass
    
    win32gui.EnumWindows(enum_callback, None)
    
    print(f"   📊 Found {len(all_pids)} total Excel PIDs: {sorted(all_pids)}")
    print(f"   🎯 Target PID: {target_pid}")
    print(f"   📋 Other PIDs to process: {sorted([pid for pid in all_pids if pid != target_pid])}")
    
    return excel_windows

def extract_filename_from_title(title: str) -> str:
    """
    Extract filename from Excel window title, handling various formats
    """
    # Remove Excel suffix
    if ' - Excel' in title:
        filename_part = title.split(' - Excel')[0]
    else:
        filename_part = title
    
    # Remove various Excel mode indicators
    indicators_to_remove = [
        ' [Protected View]',
        ' [Read-Only]', 
        ' [Compatibility Mode]',
        ' [Group]',
        ' - Saved'
    ]
    
    for indicator in indicators_to_remove:
        filename_part = filename_part.replace(indicator, '')
    
    return filename_part.strip()

def find_file_by_name(filename: str) -> str:
    """
    Try to find a file by searching common locations
    """
    if not filename.endswith(('.xlsx', '.xls', '.xlsm', '.csv')):
        return None
    
    # Search common locations
    search_locations = [
        os.path.expanduser("~"),  # User home
        os.path.join(os.path.expanduser("~"), "Documents"),
        os.path.join(os.path.expanduser("~"), "Desktop"),
        os.path.join(os.path.expanduser("~"), "Downloads"),
        "C:\\Users\\Public\\Documents",
        "C:\\temp",
        "C:\\"
    ]
    
    for location in search_locations:
        try:
            full_path = os.path.join(location, filename)
            if os.path.exists(full_path):
                return full_path
        except:
            continue
    
    return None

def extract_files_from_excel_instances(excel_windows: List[Dict]) -> List[str]:
    """
    Enhanced extraction that handles protected view and multiple PIDs
    """
    print(f"📋 Scanning {len(excel_windows)} Excel instances for open files...")
    
    # Group windows by PID to understand the process structure
    windows_by_pid = {}
    for window in excel_windows:
        pid = window['pid']
        if pid not in windows_by_pid:
            windows_by_pid[pid] = []
        windows_by_pid[pid].append(window)
    
    print(f"   📊 Windows grouped by PID:")
    for pid, windows in windows_by_pid.items():
        protected_count = sum(1 for w in windows if w['is_protected_view'])
        print(f"      PID {pid}: {len(windows)} windows ({protected_count} protected view)")
    
    # Debug: Print all windows before processing
    print(f"   🔍 DEBUG: Full list of {len(excel_windows)} windows to process:")
    for i, excel_info in enumerate(excel_windows):
        status_flags = []
        if excel_info['is_protected_view']:
            status_flags.append("PROTECTED")
        if excel_info['is_read_only']:
            status_flags.append("READ-ONLY")
        if excel_info['is_compatibility_mode']:
            status_flags.append("COMPAT")
        
        status_str = f" [{', '.join(status_flags)}]" if status_flags else ""
        print(f"      {i+1}. HWND: {excel_info['hwnd']}, PID: {excel_info['pid']}, Title: {excel_info['title']}{status_str}")
    
    all_files = set()  # Use set to avoid duplicates
    
    print("   🔍 Using enhanced window activation method...")
    
    processed_count = 0
    for i, excel_info in enumerate(excel_windows):
        processed_count += 1
        print(f"      🔍 DEBUG: Processing window {processed_count}/{len(excel_windows)} (enumerate index {i})")
        
        try:
            hwnd = excel_info['hwnd']
            pid = excel_info['pid']
            title = excel_info['title']
            is_protected = excel_info['is_protected_view']
            
            status_info = " [PROTECTED VIEW]" if is_protected else ""
            print(f"      🪟 Window {processed_count}/{len(excel_windows)}: {title}{status_info} (PID: {pid}, HWND: {hwnd})")
            
            # Verify the window still exists
            try:
                if not win32gui.IsWindow(hwnd):
                    print(f"         ❌ HWND {hwnd} is no longer a valid window - skipping")
                    continue
                
                if not win32gui.IsWindowVisible(hwnd):
                    print(f"         ❌ HWND {hwnd} is no longer visible - skipping")
                    continue
                
                print(f"         ✅ Window {hwnd} is still valid")
                
            except Exception as validation_err:
                print(f"         ❌ Window validation failed: {validation_err} - skipping")
                continue
            
            # Handle Protected View files differently
            if is_protected:
                print(f"         🔐 Handling Protected View file...")
                
                # For protected view, we usually can't access via COM
                # Try to extract filename and find the file
                filename = extract_filename_from_title(title)
                print(f"         📄 Extracted filename: {filename}")
                
                if filename.endswith(('.xlsx', '.xls', '.xlsm')):
                    file_path = find_file_by_name(filename)
                    if file_path:
                        all_files.add(file_path)
                        print(f"         ✅ Found protected view file: {file_path}")
                    else:
                        print(f"         ⚠️  Could not locate protected view file: {filename}")
                        print(f"              💡 This file may need to be manually reopened")
                
                # Don't try COM activation for protected view
                continue
            
            # For non-protected files, try window activation + COM
            activation_success = reliable_window_activation(hwnd)
            
            if not activation_success:
                print(f"         ⚠️  Could not activate window, trying anyway...")
            
            # Try to connect to Excel after activation
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                
                # Verify which Excel instance we're connected to
                try:
                    current_hwnd = excel.Hwnd
                    _, current_pid = win32process.GetWindowThreadProcessId(current_hwnd)
                    print(f"         📊 Connected to Excel PID {current_pid} (HWND: {current_hwnd})")
                    
                    # Check if we're connected to the expected process
                    if current_pid == pid:
                        print(f"         🎯 Perfect match - connected to target PID!")
                    else:
                        print(f"         ⚠️  Connected to different PID (expected {pid}, got {current_pid})")
                        print(f"              ℹ️  This is normal when multiple Excel processes exist")
                    
                except Exception as pid_err:
                    print(f"         ⚠️  Could not verify Excel PID: {pid_err}")
                
                # Get workbooks from this Excel connection
                window_files = []
                try:
                    workbook_count = excel.Workbooks.Count
                    print(f"         📚 Found {workbook_count} workbooks in this Excel instance")
                    
                    for j, wb in enumerate(excel.Workbooks, 1):
                        try:
                            wb_name = wb.Name
                            wb_path = wb.Path
                            
                            print(f"            📄 Workbook {j}: '{wb_name}' (Path: '{wb_path}')")
                            
                            if wb_path and wb_path.strip():
                                full_path = os.path.join(wb_path, wb_name)
                                if os.path.exists(full_path):
                                    window_files.append(full_path)
                                    all_files.add(full_path)
                                    print(f"               ✅ Added to restore list")
                                else:
                                    print(f"               ❌ File path doesn't exist: {full_path}")
                            else:
                                # Unsaved workbook
                                if wb_name.lower().startswith("book"):
                                    print(f"               ℹ️  Target workbook (will be processed): {wb_name}")
                                else:
                                    print(f"               ⚠️  Unsaved workbook will be lost: {wb_name}")
                                    
                        except Exception as wb_err:
                            print(f"            ❌ Error reading workbook {j}: {wb_err}")
                            continue
                    
                except Exception as workbooks_err:
                    print(f"         ❌ Error accessing workbooks: {workbooks_err}")
                
                excel_info['files'] = window_files
                print(f"         ✅ Successfully captured {len(window_files)} files from this window")
                
            except Exception as excel_err:
                print(f"         ❌ Could not connect to Excel: {excel_err}")
                
                # Fallback: Try to extract filename from window title
                print(f"         🔄 Fallback: Trying title-based file detection...")
                filename = extract_filename_from_title(title)
                
                if filename.endswith(('.xlsx', '.xls', '.xlsm')):
                    print(f"         📄 Extracted filename from title: {filename}")
                    
                    file_path = find_file_by_name(filename)
                    if file_path:
                        all_files.add(file_path)
                        print(f"         ✅ Found file via title parsing: {file_path}")
                    else:
                        print(f"         ⚠️  Could not locate file: {filename}")
            
            # Small delay between windows
            time.sleep(0.3)
            
        except Exception as window_err:
            print(f"      ❌ Error processing window {processed_count}: {window_err}")
            import traceback
            traceback.print_exc()
            continue
    
    print(f"   🔍 DEBUG: Finished processing. Processed {processed_count} out of {len(excel_windows)} windows")
    
    # Convert set back to list
    final_files = list(all_files)
    print(f"   ✅ Total unique files captured: {len(final_files)}")
    
    if final_files:
        print("   📋 Files to restore:")
        for i, file_path in enumerate(final_files, 1):
            print(f"      {i}. {os.path.basename(file_path)}")
    else:
        print("   ⚠️  No files captured for restoration")
        print("        💡 This might happen if:")
        print("           • All files are in Protected View")
        print("           • Files are unsaved")
        print("           • COM access is blocked")
    
    return final_files

def capture_open_files_before_closing(target_pid: int) -> List[str]:
    """
    Comprehensive capture of all open Excel files before closing other processes
    """
    print(f"📋 Comprehensively capturing ALL open files before closing other Excel instances...")
    
    # Get info about all Excel windows (except target)
    excel_windows = get_all_excel_window_info(target_pid)
    
    if not excel_windows:
        print("   ℹ️  No other Excel windows found")
        return []
    
    print(f"   📊 Found {len(excel_windows)} other Excel windows:")
    for i, info in enumerate(excel_windows, 1):
        print(f"      {i}. PID {info['pid']}: {info['title']}")
    
    # Extract files using multiple methods
    files_to_reopen = extract_files_from_excel_instances(excel_windows)
    
    if files_to_reopen:
        print(f"   ✅ Successfully captured {len(files_to_reopen)} files:")
        for i, file_path in enumerate(files_to_reopen, 1):
            print(f"      {i}. {os.path.basename(file_path)}")
    else:
        print("   ⚠️  No saved files found to restore")
    
    return files_to_reopen

def close_other_excel_processes(target_pid: int) -> Tuple[List[int], List[str]]:
    """
    Close all Excel processes except target, return closed PIDs and files to restore
    """
    print(f"🎯 Target Excel PID: {target_pid}")
    
    # First, capture what files are open
    files_to_reopen = capture_open_files_before_closing(target_pid)
    
    # Get all Excel processes
    all_excel_pids = get_all_excel_processes()
    other_pids = [pid for pid in all_excel_pids if pid != target_pid]
    
    if not other_pids:
        print("✅ No other Excel processes to close")
        return [], files_to_reopen
    
    print(f"🔪 Closing {len(other_pids)} other Excel processes...")
    
    closed_pids = []
    for pid in other_pids:
        try:
            print(f"   🔪 Closing Excel PID {pid}")
            psutil.Process(pid).terminate()
            closed_pids.append(pid)
            time.sleep(0.3)  # Small delay between closures
        except Exception as e:
            print(f"   ⚠️  Could not close PID {pid}: {e}")
    
    # Give processes time to close completely
    print("   ⏳ Waiting for processes to close...")
    time.sleep(2.0)
    
    print(f"✅ Closed {len(closed_pids)} Excel processes")
    return closed_pids, files_to_reopen

def restore_excel_files(files_to_reopen: List[str]):
    """
    Restore Excel files by reopening them from their file paths
    """
    if not files_to_reopen:
        print("ℹ️  No files to restore")
        return
    
    print(f"🔄 Restoring {len(files_to_reopen)} Excel files...")
    
    successfully_opened = 0
    failed_to_open = []
    
    for i, file_path in enumerate(files_to_reopen, 1):
        try:
            file_name = os.path.basename(file_path)
            print(f"   📂 {i}/{len(files_to_reopen)}: Opening {file_name}")
            
            # Verify file still exists
            if not os.path.exists(file_path):
                print(f"      ❌ File no longer exists: {file_path}")
                failed_to_open.append(file_path)
                continue
            
            # Open the file using Windows file association
            os.startfile(file_path)
            successfully_opened += 1
            
            # Small delay between file openings to avoid overwhelming Excel
            time.sleep(1.0)
            
        except Exception as e:
            print(f"      ❌ Failed to open {file_name}: {e}")
            failed_to_open.append(file_path)
    
    print(f"✅ Restoration complete:")
    print(f"   📂 Successfully opened: {successfully_opened} files")
    
    if failed_to_open:
        print(f"   ❌ Failed to open: {len(failed_to_open)} files")
        for failed_file in failed_to_open:
            print(f"      • {os.path.basename(failed_file)}")

def process_target_workbook(target_pid: int, workbook_name: str):
    """
    Process the target workbook in isolation
    """
    try:
        print(f"🔗 Connecting to isolated Excel instance (PID: {target_pid})...")
        
        # Connect to the only remaining Excel
        excel = win32com.client.GetActiveObject("Excel.Application")
        print(f"✅ Connected to Excel with {excel.Workbooks.Count} workbooks")
        
        # Find our target workbook
        target_workbook = None
        for wb in excel.Workbooks:
            try:
                wb_name = wb.Name
                wb_path = wb.Path
                print(f"   📋 Found workbook: '{wb_name}' (Path: '{wb_path}')")
                
                if (wb_name == workbook_name and 
                    wb_name.lower().startswith("book") and 
                    wb_path == ""):
                    target_workbook = wb
                    break
            except:
                continue
        
        if not target_workbook:
            print(f"❌ Target workbook '{workbook_name}' not found in isolated Excel")
            return False, []
        
        print(f"✅ Found target workbook: {target_workbook.Name}")
        
        # Process the workbook
        return save_and_process_workbook(excel, target_workbook, target_pid)
        
    except Exception as e:
        print(f"❌ Error processing isolated workbook: {e}")
        import traceback
        traceback.print_exc()
        return False, []

def save_and_process_workbook(excel_app, workbook, pid):
    """Save and process the workbook"""
    workbook_name = None
    created_reports = []
    
    try:
        # Store workbook info before operations
        workbook_name = workbook.Name
        
        # Generate filename
        timestamp = time.strftime("%m-%d-%Y_%H.%M.%S")
        filename = f"Captured_{timestamp}.xlsx"
        save_path = os.path.join(SAVE_FOLDER, filename)
        
        print(f"💾 Saving workbook '{workbook_name}' to: {save_path}")
        
        # Save using COM
        workbook.SaveAs(save_path)
        print(f"✅ Saved successfully: {save_path}")
        
        # Close the workbook gracefully
        try:
            workbook.Close(SaveChanges=False)
            print(f"📄 Closed workbook: {workbook_name}")
        except:
            print(f"📄 Workbook {workbook_name} auto-closed (normal)")
        
        # Verify file exists and has content
        if not os.path.exists(save_path) or os.path.getsize(save_path) == 0:
            raise Exception(f"File not created or empty: {save_path}")
        
        print(f"📊 File verified: {os.path.getsize(save_path)} bytes")
        
        # Process the data
        success, reports = transform_excel_file(save_path)
        
        if success:
            created_reports = reports
            
            # Kill this Excel process
            kill_excel_process(pid)
            
            # Mark as processed
            if workbook_name:
                processed_workbooks.add((workbook_name, pid))
            
            # Clean up temp file
            try:
                os.remove(save_path)
                print(f"🗑️  Deleted temporary file: {save_path}")
            except:
                pass
            
            return True, created_reports
        else:
            return False, []
            
    except Exception as e:
        print(f"⚠️  Error saving/processing workbook: {e}")
        import traceback
        traceback.print_exc()
        return False, []

def kill_excel_process(pid):
    """Kill specific Excel process"""
    try:
        process = psutil.Process(pid)
        process.terminate()
        process.wait(timeout=3)
        print(f"🔪 Excel PID {pid} terminated")
    except Exception as e:
        print(f"⚠️  Could not kill Excel PID {pid}: {e}")

def transform_excel_file(filepath):
    """Process the Excel file and create reports - returns (success, report_paths)"""
    try:
        print(f"📊 Processing data from: {filepath}")
        captured_df = pd.read_excel(filepath, dtype={"Item ID": str})
        brand_map_df = pd.read_csv(BRAND_MAP_CSV, dtype={"Item ID": str})
        df = captured_df.merge(brand_map_df, on="Item ID", how="left")
        
        df_by_cat = df.dropna(subset=["CATEGORY"]).copy()
        df_by_cat["Brand : Category"] = (
            df["Brand"].astype(str) + " : " + df["CATEGORY"].astype(str)
        )
        df_by_cat = df_by_cat.sort_values(by="Item ID")

        # Calculate profit percentages
        def calc_profit(grouped_df):
            grouped_df["Profit %"] = (
                (grouped_df["Sale Price"] - grouped_df["Unit Cost"]) / 
                grouped_df["Sale Price"] * 100
            ).round(2).astype(str) + "%"
            grouped_df.rename(columns={
                "Sale Price": "Agg Sale Price", 
                "Unit Cost": "Agg Unit Cost"
            }, inplace=True)
            return grouped_df

        # Create reports
        id_report = calc_profit(df.groupby("Item ID", as_index=False).agg({
            "Sale Price": "sum", "Unit Cost": "sum", "Item Name": "first"
        }))
        
        brand_report = calc_profit(df.groupby("Brand", as_index=False).agg({
            "Sale Price": "sum", "Unit Cost": "sum"
        }))
        
        brcat_report = calc_profit(df_by_cat.groupby("Brand : Category", as_index=False).agg({
            "Sale Price": "sum", "Unit Cost": "sum"
        }))

        # Save reports
        timestamp = time.strftime("%m-%d-%Y_%H.%M")
        paths = {
            'id': os.path.join(PROCESSED_FOLDER, f"processed_{timestamp}_id.xlsx"),
            'brand': os.path.join(PROCESSED_FOLDER, f"processed_{timestamp}_brand.xlsx"),
            'brcat': os.path.join(PROCESSED_FOLDER, f"processed_{timestamp}_brcat.xlsx")
        }
        
        id_report.to_excel(paths['id'], index=False)
        brand_report.to_excel(paths['brand'], index=False)
        brcat_report.to_excel(paths['brcat'], index=False)

        print(f"✅ Created reports:")
        report_paths = []
        for report_type, path in paths.items():
            print(f"   📄 {report_type}: {path}")
            os.startfile(path)
            report_paths.append(path)

        return True, report_paths
        
    except Exception as e:
        print(f"⚠️  Data processing failed: {e}")
        import traceback
        traceback.print_exc()
        return False, []

def watch_for_excel_workbooks():
    """Main watching loop with reliable process isolation and restoration"""
    print("👀 Watching for unsaved Excel workbooks...")
    print(f"📁 Save folder: {SAVE_FOLDER}")
    print(f"📁 Processed folder: {PROCESSED_FOLDER}")
    print("ℹ️  Method: Process isolation with reliable file restoration")
    print("✅ All previously open Excel files will be automatically restored")
    print("Press Ctrl+C to stop")
    print("─" * 70)
    
    while True:
        try:
            excel_windows = find_unsaved_excel_windows()
            
            if not excel_windows:
                time.sleep(2)
                continue
            
            for hwnd, window_title, pid in excel_windows:
                workbook_name = window_title.split(" - ")[0] if " - " in window_title else window_title
                workbook_id = (workbook_name, pid)
                
                if workbook_id in processed_workbooks:
                    continue
                
                print(f"🎯 Found new workbook: '{workbook_name}' (PID: {pid})")
                print("🚀 Starting fully automated processing with file restoration...")
                
                # Step 1: Close other Excel processes (capturing files first)
                closed_pids, files_to_reopen = close_other_excel_processes(pid)
                
                try:
                    # Step 2: Process the target workbook in isolation
                    success, reports = process_target_workbook(pid, workbook_name)
                    
                    if success:
                        print("✅ Successfully processed workbook")
                        print(f"📋 Created {len(reports)} reports")
                    else:
                        print("⚠️  Failed to process workbook")
                
                finally:
                    # Step 3: Restore all previously open files
                    print("🔄 Restoring previously open Excel files...")
                    restore_excel_files(files_to_reopen)
                    
                    if files_to_reopen:
                        print(f"✅ Restoration process complete!")
                        print(f"   📂 Attempted to restore {len(files_to_reopen)} files")
                    else:
                        print("ℹ️  No files needed restoration")
                
                print("─" * 70)
                break  # Process one workbook at a time
                
        except Exception as e:
            print(f"⚠️  Error in watch loop: {e}")
            import traceback
            traceback.print_exc()
        
        time.sleep(2)

def create_test_excel_processes():
    """
    Helper function to create multiple Excel processes for testing
    """
    print("🧪 Creating multiple Excel processes for testing...")
    
    methods = [
        "1. File → New Window (from existing Excel)",
        "2. Right-click Excel taskbar icon → Excel", 
        "3. Command line: excel.exe /x",
        "4. Open files from different security contexts",
        "5. Use this automated method"
    ]
    
    print("📋 Ways to create multiple Excel processes:")
    for method in methods:
        print(f"   {method}")
    
    try:
        # Method: Start Excel with /x flag (forces new instance)
        print("\n🚀 Attempting to start new Excel instances...")
        
        import subprocess
        
        # Try to find Excel executable
        excel_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE", 
            r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE",
        ]
        
        excel_exe = None
        for path in excel_paths:
            if os.path.exists(path):
                excel_exe = path
                break
        
        if not excel_exe:
            print("   ❌ Could not find Excel executable")
            print("   💡 Try manual methods above")
            return
        
        print(f"   ✅ Found Excel at: {excel_exe}")
        
        # Start 2 new Excel instances
        for i in range(2):
            try:
                subprocess.Popen([excel_exe, "/x"], 
                               creationflags=subprocess.CREATE_NEW_PROCESS_GROUP)
                print(f"   🚀 Started Excel instance {i+1}")
                time.sleep(2)  # Wait between launches
            except Exception as e:
                print(f"   ❌ Failed to start instance {i+1}: {e}")
        
        print("\n📊 Current Excel processes:")
        excel_pids = get_all_excel_processes()
        for i, pid in enumerate(excel_pids, 1):
            print(f"   {i}. PID: {pid}")
        
        print(f"\n✅ Now you have {len(excel_pids)} Excel processes for testing!")
        print("💡 Open different files in each instance to test multi-PID restoration")
        
    except Exception as e:
        print(f"❌ Automated method failed: {e}")
        print("💡 Please use manual methods above")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1 and sys.argv[1] == "--test-setup":
        print("🧪 Excel Multi-Process Test Setup")
        print("=" * 50)
        create_test_excel_processes()
        sys.exit(0)
    
    print("🚀 Excel Automation - Reliable Process Isolation")
    print("=" * 70)
    print("✅ Zero user interaction during processing")
    print("✅ Handles Protected View files")
    print("✅ Supports multiple Excel processes")
    print("✅ Enhanced file path detection")
    print("=" * 70)
    print("💡 To test multiple processes: python script.py --test-setup")
    print("=" * 70)
    
    try:
        watch_for_excel_workbooks()
    except KeyboardInterrupt:
        print("\n👋 Excel watcher stopped by user")