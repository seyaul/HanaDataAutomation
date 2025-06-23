import os, sys, textwrap, pythoncom, win32com.client as win32

# ─── settings ─────────────────────────────────────────────────────────────
CAPTURE_DIR   = r"C:\Users\sasuk\Documents\CapturedExports"
MODULE_CLASS  = "cAppEvents"          # class module name
MODULE_STD    = "AutoSaveBootstrap"   # standard module name
XLSTART_PATH  = os.path.expandvars(r"%appdata%\Microsoft\Excel\XLStart")
PERSONAL_XLSB = os.path.join(XLSTART_PATH, "PERSONAL.XLSB")
# ──────────────────────────────────────────────────────────────────────────

CLASS_CODE = textwrap.dedent(f"""
    ''# {MODULE_CLASS}.cls  – compiled into PERSONAL.XLSB
    Option Explicit
    Private WithEvents App As Excel.Application

    Private Sub Class_Initialize()
        Set App = Application
    End Sub

    Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
        If Wb.Path = "" And LCase$(Left$(Wb.Name, 4)) = "book" Then
            Dim tgt As String
            tgt = "{CAPTURE_DIR}" & "\\Captured_" & _
                  Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
            If Dir$(\"{CAPTURE_DIR}\", vbDirectory) = \"\" _
                    Then MkDir \"{CAPTURE_DIR}\"
            Application.DisplayAlerts = False
            Wb.SaveAs tgt, xlOpenXMLWorkbook
            Application.DisplayAlerts = True
        End If
    End Sub
""").lstrip()

STD_CODE = textwrap.dedent(f"""
    ''# {MODULE_STD}.bas – creates global event sink on startup
    Option Explicit
    Public gAppEvents As {MODULE_CLASS}

    Sub Auto_Open()
        Set gAppEvents = New {MODULE_CLASS}
    End Sub
""").lstrip()

# ──────────────────────────────────────────────────────────────────────────
def ensure_personal_exists(excel):
    """Create PERSONAL.XLSB if the user never recorded a macro."""
    if not os.path.exists(PERSONAL_XLSB):
        wb = excel.Workbooks.Add()
        excel.DisplayAlerts = False
        wb.SaveAs(Filename=PERSONAL_XLSB, FileFormat=50)   # 50 = xlExcel12 (xlsb)
        wb.Close()
        excel.DisplayAlerts = True
        print("➕ Created blank PERSONAL.XLSB")

def add_or_replace_component(vbproj, name, code, kind):
    """
    kind: 1=Standard, 2=Class
    """
    # remove old
    for comp in vbproj.VBComponents:
        if comp.Name == name:
            vbproj.VBComponents.Remove(comp)
            break
    comp = vbproj.VBComponents.Add(kind)
    comp.Name = name
    comp.CodeModule.AddFromString(code)
    print(f"✓ injected {name}")

def main():
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False

    ensure_personal_exists(excel)
    wb = excel.Workbooks.Open(PERSONAL_XLSB)
    vbproj = wb.VBProject

    add_or_replace_component(vbproj, MODULE_CLASS, CLASS_CODE, 2)  # 2 = class module
    add_or_replace_component(vbproj, MODULE_STD,   STD_CODE,   1)  # 1 = standard

    wb.Save()
    excel.Quit()
    print("\n✅ Auto-save VBA installed.  Restart Excel to activate.")

if __name__ == "__main__":
    main()