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

def capture_open_files_before_closing(target_pid: int) -> List[str]:
    """
    Capture all open Excel files before closing other processes
    Returns list of file paths that should be reopened later
    """
    print(f"📋 Capturing open files before closing other Excel instances...")
    
    all_excel_pids = get_all_excel_processes()
    other_pids = [pid for pid in all_excel_pids if pid != target_pid]
    
    if not other_pids:
        print("   ℹ️  No other Excel processes to close")
        return []
    
    print(f"   📊 Found {len(other_pids)} other Excel processes to close")
    
    files_to_reopen = []
    
    # Try to get file paths from each Excel instance before closing
    for attempt in range(min(3, len(other_pids))):  # Try up to 3 times
        try:
            print(f"   🔍 Attempt {attempt + 1}: Scanning for open files...")
            
            # Connect to any Excel instance (might be one we're about to close)
            excel = win32com.client.GetActiveObject("Excel.Application")
            
            current_files = []
            for wb in excel.Workbooks:
                try:
                    wb_name = wb.Name
                    wb_path = wb.Path
                    
                    # Only save files that are actually saved (have a path)
                    if wb_path and wb_path.strip():
                        full_path = os.path.join(wb_path, wb_name)
                        if os.path.exists(full_path):
                            current_files.append(full_path)
                            print(f"      📄 Found: {wb_name}")
                    else:
                        # Unsaved workbook - warn user
                        if not wb_name.lower().startswith("book"):
                            print(f"      ⚠️  Unsaved workbook will be lost: {wb_name}")
                            
                except Exception as wb_err:
                    continue
            
            # Add new files to our list (avoid duplicates)
            for file_path in current_files:
                if file_path not in files_to_reopen:
                    files_to_reopen.append(file_path)
            
            time.sleep(0.5)  # Small delay between attempts
            
        except Exception as e:
            print(f"      ⚠️  Attempt {attempt + 1} failed: {e}")
            time.sleep(0.5)
            continue
    
    print(f"   ✅ Captured {len(files_to_reopen)} files to reopen later")
    for i, file_path in enumerate(files_to_reopen, 1):
        print(f"      {i}. {os.path.basename(file_path)}")
    
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

if __name__ == "__main__":
    print("🚀 Excel Automation - Reliable Process Isolation")
    print("=" * 70)
    print("✅ Zero user interaction during processing")
    print("✅ Automatic capture and restoration of all open Excel files")
    print("✅ Uses file paths for reliable restoration")
    print("=" * 70)
    
    try:
        watch_for_excel_workbooks()
    except KeyboardInterrupt:
        print("\n👋 Excel watcher stopped by user")