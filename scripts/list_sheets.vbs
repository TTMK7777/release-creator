Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Open("C:\Users\t-tsuji\AIアプリ開発\release-creator\テンプレート\【テンプレ】リリース内表 - コピー.xlsx")

WScript.Echo "=== ターゲットファイルのシート名一覧 ==="
For Each objWorksheet in objWorkbook.Worksheets
    WScript.Echo objWorksheet.Name
Next

objWorkbook.Close False
objExcel.Quit

Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
