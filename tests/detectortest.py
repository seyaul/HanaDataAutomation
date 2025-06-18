import win32com.client
import time

print("🕵️‍♂️ Excel workbook detector running...")

known_workbooks = set()

while True:
    try:
        # Try to connect to running Excel instance
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            print("⚠️ Excel not open. Waiting...")
            time.sleep(2)
            continue

        # Loop through currently open workbooks
        for wb in excel.Workbooks:
            wb_id = (wb.Name, wb.Path)
            if wb_id not in known_workbooks:
                print(f"🆕 New workbook detected: {wb.Name}, Path: '{wb.Path}'")
                known_workbooks.add(wb_id)

        # Optionally list current workbooks every loop
        # print(f"Open workbooks: {[wb.Name for wb in excel.Workbooks]}")
    except Exception as e:
        print(f"❌ Error: {e}")

    time.sleep(2)
