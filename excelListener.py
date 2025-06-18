import win32com.client
import os
import time
import pandas as pd

SAVE_FOLDER = r"C:\Users\sasuk\Documents\CapturedExports"
PROCESSED_FOLDER = r"C:\Users\sasuk\Documents\ProcessedExports"

os.makedirs(SAVE_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def transform_excel_file(filepath):
    try:
        df = pd.read_excel(filepath)

        # 🛠️ Your transformation logic
        df.dropna(how="all", inplace=True)
        df['Double First Column'] = df.iloc[:, 0].apply(lambda x: x * 2 if pd.notna(x) else None)

        processed_filename = "processed_" + os.path.basename(filepath)
        processed_path = os.path.join(PROCESSED_FOLDER, processed_filename)
        df.to_excel(processed_path, index=False)

        print(f"✅ Transformed and saved: {processed_path}")
        os.startfile(processed_path)  # Open processed file
        #os.remove(filepath)  # Delete captured temp file
        print(f"🗑️ Deleted captured file: {filepath}")
    except Exception as e:
        print(f"⚠️ Error during processing: {e}")

def auto_capture_and_transform():
    known_workbooks = set()
    print("👀 Watching Excel for new workbooks...")

    while True:
        try:
            # Try to connect to Excel instance
            try:
                excel = get_excel_instance()
                if not excel:
                    raise Exception
                for wb in excel.Workbooks:
                    print(wb.Name, " ", wb.Path)
            except Exception:
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

                    wb.SaveAs(save_path)
                    print(f"💾 Saved: {save_path}")
                    #wb.Close(SaveChanges=False)
                    transform_excel_file(save_path)

                    known_workbooks.add(wb_id)

        except Exception as e:
            print(f"⚠️ Error in detection loop: {e}")
        time.sleep(2)


def get_excel_instance():
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        if excel.Workbooks.Count > 0:
            return excel
        else:
            print("⚠️ Excel is running but has no open workbooks.")
            return None
    except Exception:
        print("⚠️ Excel not detected at all.")
        return None

if __name__ == "__main__":
    auto_capture_and_transform()
