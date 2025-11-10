Attribute VB_Name = "Module2_Data"

Public Sub TransferData(dataFilePath As String)
    On Error GoTo ErrorHandler
    
    ' データ転記処理
    ' TODO: データの読み込みと転記ロジックを実装

    Exit Sub
    
ErrorHandler:
    MsgBox "データ転記中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
