Attribute VB_Name = "Module4_Word"
Option Explicit

Sub UpdateWordDocument()
    On Error GoTo ErrorHandler
    
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim templatePath As String
    Dim savePath As String
    Dim findText As String
    Dim replaceText As String
    
    ' Define paths
    templatePath = "C:\path\to\template.docx"
    savePath = "C:\path\to\updated_document.docx"
    
    ' Open Word application
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Set wdDoc = wdApp.Documents.Open(templatePath)
    
    ' Update date
    findText = "2024"
    replaceText = "2025"
    With wdDoc.Content.Find
        .Text = findText
        .Replacement.Text = replaceText
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    ' Update publication date
    Call UpdatePublicationDate(wdDoc)
    
    ' Update title
    Call UpdateTitle(wdDoc)
    
    ' Replace image
    Call ReplaceImage(wdDoc)
    
    ' Save the updated document
    wdDoc.SaveAs2 FileName:=savePath
    wdDoc.Close
    wdApp.Quit
    
    ' Log success
    Call Module1_Main.LogMessage("Word document updated successfully.")
    Exit Sub

ErrorHandler:
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Call Module1_Main.LogMessage("Error updating Word document: " & Err.Description)
End Sub

Sub UpdatePublicationDate(wdDoc As Object)
    ' Example implementation to update publication date - replace with actual logic
    Dim findText As String
    Dim replaceText As String
    findText = "Publication Date: XXXX"
    replaceText = "Publication Date: " & Format(Date, "yyyy-mm-dd")
    With wdDoc.Content.Find
        .Text = findText
        .Replacement.Text = replaceText
        .Execute Replace:=2 ' wdReplaceAll
    End With
End Sub

Sub UpdateTitle(wdDoc As Object)
    ' Example implementation to update title - replace with actual logic
    Dim findText As String
    Dim replaceText As String
    findText = "Old Title"
    replaceText = "New Title"
    With wdDoc.Content.Find
        .Text = findText
        .Replacement.Text = replaceText
        .Execute Replace:=2 ' wdReplaceAll
    End With
End Sub

Sub ReplaceImage(wdDoc As Object)
    ' Example implementation to replace an image in the document
    Dim shape As Object
    For Each shape In wdDoc.InlineShapes
        If shape.Type = 3 Then ' wdInlineShapePicture
            shape.Delete
            wdDoc.InlineShapes.AddPicture FileName:="C:\path\to\new_image.png", _
                                          LinkToFile:=False, SaveWithDocument:=True
            Exit For
        End If
    Next shape
End Sub
