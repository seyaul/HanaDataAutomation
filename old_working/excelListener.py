import win32com.client
import os
import time
import pandas as pd
import win32process
import psutil

SAVE_FOLDER = r"C:\Users\sasuk\Documents\CapturedExports"
PROCESSED_FOLDER = r"C:\Users\sasuk\Documents\ProcessedExports"

os.makedirs(SAVE_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)
output_csv = "test123.csv"
known_workbooks = set()

def get_excel_pid_from_hwnd(hwnd):
    """Given a window handle, return the Excel process ID"""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception as e:
        print(f"⚠️ Failed to get PID: {e}")
        return None

def kill_pid(pid):
    try:
        p = psutil.Process(pid)
        p.terminate()
        print(f"Killed excel process with PID: {pid}")
    except Exception as e:
        print(f"Failed to kill excel process {pid} eith error: {e}")


def transform_excel_file(filepath):
    try:
        captured_df = pd.read_excel(filepath, dtype={"Item ID": str})
        brand_map_df = pd.read_csv("brand_map.csv", dtype={"Item ID": str})
        df = captured_df.merge(brand_map_df, on="Item ID", how="left")
        df_by_cat = df.dropna(subset=["CATEGORY"]).copy()
        df_by_cat["Brand : Category"] = df["Brand"].astype(str) + " : " + df["CATEGORY"].astype(str)
        df_by_cat = df_by_cat.sort_values(by="Item ID")

        # print("=" * 50)
        # print(df_by_cat.head(10))
        #df_by_cat.to_csv(output_csv, index=False)

        # TODO: Create helper function instead
        grouped_df_id = calc_profit_percentage_accname(df,0)
        grouped_df_br = calc_profit_percentage_brand(df,0)
        grouped_df_brcat = calc_profit_percentage_brand(df_by_cat, 1)

        processed_filename_id = "processed_ver-id_" + os.path.basename(filepath) 
        processed_filename_br = "processed_ver-br_" + os.path.basename(filepath) 
        processed_filename_brcat = "processed_ver-brcat_" + os.path.basename(filepath) 
        processed_path_id = os.path.join(PROCESSED_FOLDER, processed_filename_id)
        processed_path_br = os.path.join(PROCESSED_FOLDER, processed_filename_br)
        processed_path_brcat = os.path.join(PROCESSED_FOLDER, processed_filename_brcat)
        grouped_df_id.to_excel(processed_path_id, index=False)
        grouped_df_br.to_excel(processed_path_br, index=False)
        grouped_df_brcat.to_excel(processed_path_brcat, index=False)

        print(f"✅ Transformed and saved: {processed_path_id}, {processed_path_br}, and {processed_path_brcat}")
        os.startfile(processed_path_id)  # Open processed file
        os.startfile(processed_path_br)
        os.startfile(processed_path_brcat)

        # Cleanup for storage purposes. Feel free to remove
        # os.remove(filepath)  # Delete captured temp file
        # print(f"🗑️ Deleted captured file: {filepath}")
        return True
    except Exception as e:
        print(f"⚠️ Error during processing: {e}")
        return False

def auto_capture_and_transform():

    print("👀 Watching Excel for new workbooks...")
    last_failed = True
    while True:
        try:
            # Try to connect to Excel instance
            try:
                excel = get_excel_instance()
                if not excel:
                    raise Exception
                print("Checking try-catch connecting to excel instance.. being here means there is a different primary pid, or excel is running" \
                "in the background")
            except Exception:
                if last_failed:
                    print("⚠️Excel not found. trying again.")
                time.sleep(5)
                continue

            # Detect new workbooks
            for wb in excel.Workbooks:
                wb_id = (wb.Name, wb.Path)
                if wb_id not in known_workbooks and wb.Path == "" :
                    print(f"📄 New workbook detected: {wb.Name}, Path: '{wb.Path}'")

                    timestamp = time.strftime("%m-%d-%Y_%H.%M")
                    filename = f"Captured_{timestamp}.xlsx"
                    save_path = os.path.join(SAVE_FOLDER, filename)
                    
                    hwnd = excel.Hwnd
                    pid = get_excel_pid_from_hwnd(hwnd)
                    print(pid)

                    wb.SaveAs(save_path)
                    print(f"💾 Saved: {save_path}")
                    wb.Close(SaveChanges=False)
                    success = transform_excel_file(save_path)
                    print(pid)
                    if success and pid:
                        kill_pid(pid)
                        break
                    else:
                        raise Exception
                    # known_workbooks.add(wb_id)
       

        except Exception as e:
            if last_failed:
                print(f"⚠️ Error in detection loop: {e}")
                last_failed = False
        print("Repeating the while loop!")
        time.sleep(2)


def get_excel_instance():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        if excel.Workbooks.Count > 0:
            return excel
        else:
            print("1")
            return None
    except Exception:
        print("2")
        return None

def calc_profit_percentage_accname(df, vernum):
    if vernum == 0:
        grouped_df = df.groupby("Account Name", as_index=False).agg({
            "Sale Price": "sum",
            "Unit Cost": "sum",
        })
        grouped_df["Profit %"] = ((grouped_df["Sale Price"] - grouped_df["Unit Cost"])/ grouped_df["Sale Price"]) * 100
        grouped_df["Profit %"] = grouped_df["Profit %"].round(2).astype(str) + "%"
        grouped_df.rename(columns = {"Sale Price": "Agg Sale Price", "Unit Cost": "Agg Unit Cost"}, inplace = True)
        return grouped_df

def calc_profit_percentage_brand(df, vernum):
    if vernum == 0:
        grouped_df = df.groupby("Brand", as_index=False).agg({
            "Sale Price": "sum",
            "Unit Cost": "sum",
        })
        grouped_df["Profit %"] = ((grouped_df["Sale Price"] - grouped_df["Unit Cost"])/ grouped_df["Sale Price"]) * 100
        grouped_df["Profit %"] = grouped_df["Profit %"].round(2).astype(str) + "%"
        grouped_df.rename(columns = {"Sale Price": "Agg Sale Price", "Unit Cost": "Agg Unit Cost"}, inplace = True)
        return grouped_df
    if vernum == 1:
        grouped_df = df.groupby("Brand : Category", as_index = False).agg({
            "Sale Price": "sum",
            "Unit Cost": "sum",
        })
        grouped_df["Profit %"] = ((grouped_df["Sale Price"] - grouped_df["Unit Cost"])/ grouped_df["Sale Price"]) * 100
        grouped_df["Profit %"] = grouped_df["Profit %"].round(2).astype(str) + "%"
        grouped_df.rename(columns = {"Sale Price": "Agg Sale Price", "Unit Cost": "Agg Unit Cost"}, inplace = True)
        return grouped_df


if __name__ == "__main__":
    auto_capture_and_transform()