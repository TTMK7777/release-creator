Attribute VB_Name = "Module5_Utils"

Public Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

Public Sub HandleError(errorMessage As String)
    MsgBox errorMessage, vbCritical
End Sub
