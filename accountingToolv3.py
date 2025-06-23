import os, time, ctypes, pythoncom, win32gui, win32process
import win32com.client, pandas as pd, psutil


# ── folders ───────────────────────────────────────────────────────────────
CAPTURE_DIR   = r"C:\Users\sasuk\Documents\CapturedExports"
PROCESSED_DIR = r"C:\Users\sasuk\Documents\ProcessedExports"
BRAND_MAP_CSV = "brand_map.csv"

os.makedirs(CAPTURE_DIR,   exist_ok=True)
os.makedirs(PROCESSED_DIR, exist_ok=True)

# ── HWND → Excel.Application helper (uses Accessibility) ──────────────────
oleacc = ctypes.oledll.oleacc
OBJID_NATIVEOM = 0xFFFFFFF0   # defined in winuser.h

class GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]

def app_from_hwnd(hwnd):
    """
    Return Excel.Application COM object for an XLMAIN hwnd, or None
    if the window doesn't expose the native object model.
    """
    pythoncom.CoInitialize()

    # build GUID for IID_IDispatch
    iid = GUID.from_buffer_copy(bytes(pythoncom.IID_IDispatch))

    pdisp = ctypes.c_void_p()
    try:
        hr = oleacc.AccessibleObjectFromWindow(
            hwnd,
            OBJID_NATIVEOM,
            ctypes.byref(iid),
            ctypes.byref(pdisp)
        )
        if hr != 0 or not pdisp:
            return None
        return win32com.client.Dispatch(pdisp.value)
    except OSError:
        # E_FAIL or other COM errors: window isn't ready or supported
        return None

# ── transformation logic (keep/modify as needed) ──────────────────────────
def calc_profit(df, by_cols):
    grouped = (
        df.groupby(by_cols, as_index=False)
          .agg({"Sale Price": "sum", "Unit Cost": "sum"})
    )
    grouped["Profit %"] = (
        ((grouped["Sale Price"] - grouped["Unit Cost"])
         / grouped["Sale Price"]) * 100
    ).round(2).astype(str) + "%"
    grouped.rename(columns={"Sale Price": "Agg Sale Price",
                            "Unit Cost":  "Agg Unit Cost"}, inplace=True)
    return grouped

def transform(path):
    cap   = pd.read_excel(path, dtype={"Item ID": str})
    brand = pd.read_csv(BRAND_MAP_CSV, dtype={"Item ID": str})

    df = cap.merge(brand, on="Item ID", how="left")

    df_cat = df.dropna(subset=["CATEGORY"]).copy()
    df_cat["Brand : Category"] = df["Brand"].astype(str) + " : " + df["CATEGORY"].astype(str)

    reports = {
        "by_account": calc_profit(df,       ["Account Name"]),
        "by_brand"  : calc_profit(df,       ["Brand"]),
        "by_brcat"  : calc_profit(df_cat,   ["Brand : Category"])
    }

    ts = time.strftime("%Y%m%d_%H%M%S")
    for key, frame in reports.items():
        out = os.path.join(PROCESSED_DIR, f"{key}_{ts}.xlsx")
        frame.to_excel(out, index=False)
        print("   ↳ wrote", out)

# ── watcher loop ──────────────────────────────────────────────────────────
processed_hwnds = set()
print("👀  Watching every visible Excel window …  (Ctrl+C to stop)")
try:
    while True:
        def _enum(hwnd, _):
            if win32gui.GetClassName(hwnd) != "XLMAIN":
                return
            if hwnd in processed_hwnds:
                return
            title = win32gui.GetWindowText(hwnd)
            if title.lower().startswith("book"):      # Book1 / Book2 …
                app = app_from_hwnd(hwnd)
                if not app:
                    return
                wb = app.ActiveWorkbook
                if wb and wb.Path == "":
                    ts   = time.strftime("%m-%d-%Y_%H.%M.%S")
                    save = os.path.join(CAPTURE_DIR, f"Captured_{ts}.xlsx")
                    try:
                        wb.SaveAs(save)               # COM SaveAs
                        print(f"💾 captured {title}  →  {save}")
                        wb.Close(SaveChanges=False)    # close just this tab
                        transform(save)                # process & export
                        os.remove(save)                # tidy up
                        processed_hwnds.add(hwnd)
                    except Exception as e:
                        print("⚠️  Save/transform error:", e)
        pythoncom.CoInitialize()
        win32gui.EnumWindows(_enum, None)
        time.sleep(1)
except KeyboardInterrupt:
    print("\nStopped.")
