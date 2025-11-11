Attribute VB_Name = "Module3_Image"
Option Explicit

Sub ExportRangeToImage()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim rng As Range
    Dim chartObj As ChartObject
    Dim filePath As String
    
    ' Set the worksheet and range
    Set ws = ThisWorkbook.Sheets("Ranking")
    Set rng = ws.Range("A1:D10") ' Example range, replace with actual range
    
    ' Create a temporary chart
    Set chartObj = ws.ChartObjects.Add(Left:=rng.Left, Width:=rng.Width, _
                                       Top:=rng.Top, Height:=rng.Height)
    With chartObj.Chart
        .SetSourceData Source:=rng
        .Export Filename:=ThisWorkbook.Path & "\RankingChart.png", FilterName:="PNG"
    End With
    
    ' Delete the temporary chart
    chartObj.Delete
    
    ' Log success
    Call Module1_Main.LogMessage("Exported range to image successfully.")
    Exit Sub

ErrorHandler:
    Call Module1_Main.LogMessage("Error exporting range to image: " & Err.Description)
End Sub
