Attribute VB_Name = "Module3_Image_Improved"
'========================================
' Module3_Image_Fixed
' Excel範囲を正しく画像化するモジュール
'
' 作成日: 2025-11-11
' バージョン: 3.0 (修正版 - 真っ白画像の問題を解決)
'========================================
Option Explicit

'========================================
' 公開関数: 範囲を画像として出力（修正版）
'========================================
Public Function ExportRangeToImage( _
    Optional sourceFilePath As String = "", _
    Optional sourceSheetName As String, _
    Optional sourceRange As String, _
    Optional outputFolder As String, _
    Optional outputFileName As String _
) As Boolean

    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim outputPath As String
    Dim startTime As Double
    Dim needClose As Boolean
    Dim tempChart As Chart
    Dim shp As Shape

    startTime = Timer
    needClose = False

    Module1_Main.LogMessage "画像生成を開始: " & sourceSheetName & "!" & sourceRange

    ' パラメータ検証
    If Trim(sourceSheetName) = "" Or Trim(sourceRange) = "" Or _
       Trim(outputFolder) = "" Or Trim(outputFileName) = "" Then
        Module1_Main.LogMessage "  [ERROR] パラメータが不足しています"
        ExportRangeToImage = False
        Exit Function
    End If

    ' ファイルを開く
    If sourceFilePath <> "" Then
        If Not IsFileOpen(sourceFilePath) Then
            Set wb = Workbooks.Open(sourceFilePath, ReadOnly:=True, UpdateLinks:=False)
            needClose = True
        Else
            Set wb = GetOpenWorkbook(sourceFilePath)
        End If
    Else
        Set wb = ThisWorkbook
    End If

    ' シート名の柔軟な検索（スペース対応）
    On Error Resume Next
    Set ws = wb.Worksheets(sourceSheetName)
    If ws Is Nothing Then Set ws = wb.Worksheets(Trim(sourceSheetName))
    If ws Is Nothing Then Set ws = wb.Worksheets(" " & sourceSheetName)
    If ws Is Nothing Then Set ws = wb.Worksheets(sourceSheetName & " ")
    If ws Is Nothing Then Set ws = wb.Worksheets(" " & sourceSheetName & " ")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Module1_Main.LogMessage "  [ERROR] シート「" & sourceSheetName & "」が見つかりません"
        If needClose Then wb.Close SaveChanges:=False
        ExportRangeToImage = False
        Exit Function
    End If

    ' 範囲取得
    Set rng = ws.Range(sourceRange)

    ' 出力パス生成
    outputPath = outputFolder
    If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"

    ' フォルダ作成（存在しない場合）
    If Dir(outputPath, vbDirectory) = "" Then
        MkDir outputPath
    End If

    outputPath = outputPath & outputFileName

    ' ===================================
    ' 修正版: 範囲をコピーして図として貼り付け
    ' ===================================

    ' 範囲をコピー（図としてコピー）
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture

    ' 同じシートに一時的に貼り付け
    Set shp = ws.Shapes.PasteSpecial(DataType:=xlPasteMetafilePicture)

    ' 図をPNGとして保存
    SaveShapeAsPNG shp, outputPath

    ' 一時図形を削除
    shp.Delete

    ' ファイルを閉じる
    If needClose Then
        wb.Close SaveChanges:=False
    End If

    Module1_Main.LogMessage "  [OK] 画像生成完了: " & outputPath & " (" & Format(Timer - startTime, "0.00") & "秒)"
    ExportRangeToImage = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "  [ERROR] 画像生成失敗: " & Err.Description & " (Err#" & Err.Number & ")"

    On Error Resume Next
    If Not shp Is Nothing Then shp.Delete
    If needClose And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0

    ExportRangeToImage = False
End Function

'========================================
' 図形をPNGとして保存
'========================================
Private Sub SaveShapeAsPNG(shp As Shape, outputPath As String)
    On Error GoTo ErrorHandler

    ' 一時的なChartを作成して図をPNG出力
    Dim cht As Chart
    Set cht = Charts.Add
    cht.ChartArea.Select

    ' 図をChartに貼り付け
    shp.Copy
    cht.Paste

    ' PNG出力
    cht.Export Filename:=outputPath, FilterName:="PNG"

    ' Chart削除
    cht.Delete

    Exit Sub

ErrorHandler:
    Module1_Main.LogMessage "  [ERROR] PNG保存エラー: " & Err.Description
    On Error Resume Next
    If Not cht Is Nothing Then cht.Delete
    On Error GoTo 0
End Sub

'========================================
' ファイルが既に開かれているか確認
'========================================
Private Function IsFileOpen(filePath As String) As Boolean
    Dim wb As Workbook
    Dim fileName As String

    fileName = Dir(filePath)

    On Error Resume Next
    For Each wb In Workbooks
        If wb.Name = fileName Then
            IsFileOpen = True
            Exit Function
        End If
    Next wb
    On Error GoTo 0

    IsFileOpen = False
End Function

'========================================
' 開かれているワークブックを取得
'========================================
Private Function GetOpenWorkbook(filePath As String) As Workbook
    Dim wb As Workbook
    Dim fileName As String

    fileName = Dir(filePath)

    On Error Resume Next
    For Each wb In Workbooks
        If wb.Name = fileName Then
            Set GetOpenWorkbook = wb
            Exit Function
        End If
    Next wb
    On Error GoTo 0

    Set GetOpenWorkbook = Nothing
End Function

'========================================
' 3種類の画像を生成する統合関数
'========================================
Public Function GenerateAllRankingImages( _
    targetFilePath As String, _
    outputFolder As String _
) As Boolean

    On Error GoTo ErrorHandler
    GenerateAllRankingImages = False

    Dim imageCount As Long
    imageCount = 0

    Module1_Main.LogMessage "[画像生成] 3種類のランキング画像を生成します"
    Module1_Main.LogMessage ""

    ' =========================================
    ' 1. 総合ランキング画像（総合1つシート）
    ' =========================================
    Module1_Main.LogMessage "  [1/3] 総合ランキング画像を生成中..."
    If ExportRangeToImage( _
        sourceFilePath:=targetFilePath, _
        sourceSheetName:="総合1つ", _
        sourceRange:="C4:E10", _
        outputFolder:=outputFolder, _
        outputFileName:="01_総合ランキング.png" _
    ) Then
        imageCount = imageCount + 1
    End If
    Module1_Main.LogMessage ""

    ' =========================================
    ' 2. 評価項目ランキング画像（総合2つシート右側）
    ' =========================================
    Module1_Main.LogMessage "  [2/3] 評価項目ランキング画像を生成中..."
    If ExportRangeToImage( _
        sourceFilePath:=targetFilePath, _
        sourceSheetName:="総合2つ", _
        sourceRange:="G4:I10", _
        outputFolder:=outputFolder, _
        outputFileName:="02_評価項目ランキング.png" _
    ) Then
        imageCount = imageCount + 1
    End If
    Module1_Main.LogMessage ""

    ' =========================================
    ' 3. 評価・部門別ランキング画像（評価・部門別シート）
    ' =========================================
    Module1_Main.LogMessage "  [3/3] 評価・部門別ランキング画像を生成中..."
    If ExportRangeToImage( _
        sourceFilePath:=targetFilePath, _
        sourceSheetName:=" 評価・部門別 ", _
        sourceRange:="B7:P18", _
        outputFolder:=outputFolder, _
        outputFileName:="03_評価部門別ランキング.png" _
    ) Then
        imageCount = imageCount + 1
    End If
    Module1_Main.LogMessage ""

    Module1_Main.LogMessage "[画像生成] 完了: " & imageCount & "/3 件"

    GenerateAllRankingImages = (imageCount = 3)
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "[ERROR] 画像生成処理エラー: " & Err.Description
    GenerateAllRankingImages = False
End Function
