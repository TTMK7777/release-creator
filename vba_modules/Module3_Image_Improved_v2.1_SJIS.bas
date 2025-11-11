Attribute VB_Name = "Module3_Image_Improved"
'========================================
' Module3_Image_Improved
' Excel範囲を画像化するモジュール（改善版）
'
' 作成日: 2025-11-11
' 更新日: 2025-11-11
' 作成者: GPT-4o提案 + Claude Code実装
' バージョン: 2.1 (外部ファイル対応)
'========================================
Option Explicit

'========================================
' 公開関数: 範囲を画像として出力（外部ファイル対応）
'========================================
' @param sourceFilePath   ソースExcelファイルパス（省略時はThisWorkbook）
' @param sourceSheetName  ソースシート名
' @param sourceRange      ソース範囲（例: "B7:D15"）
' @param outputFolder     出力フォルダパス
' @param outputFileName   出力ファイル名（例: "RankingChart.png"）
' @return Boolean         成功時True、失敗時False
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
    Dim chartObj As ChartObject
    Dim outputPath As String
    Dim startTime As Double
    Dim needClose As Boolean

    startTime = Timer
    needClose = False

    Call Module1_Main.LogMessage("画像生成を開始: " & sourceSheetName & "!" & sourceRange)

    ' パラメータ検証
    If Not ValidateParameters(sourceFilePath, sourceSheetName, sourceRange, outputFolder, outputFileName) Then
        ExportRangeToImage = False
        Exit Function
    End If

    ' ファイルを開く（外部ファイルの場合）
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

    ' シート取得
    Set ws = wb.Worksheets(sourceSheetName)

    ' 範囲取得
    Set rng = ws.Range(sourceRange)

    ' 出力パス生成
    outputPath = outputFolder
    If Right(outputPath, 1) <> "\" Then outputPath = outputPath & "\"
    outputPath = outputPath & outputFileName

    ' 一時的なChartObjectを作成
    Set chartObj = ws.ChartObjects.Add( _
        Left:=rng.Left, _
        Top:=rng.Top, _
        Width:=rng.Width, _
        Height:=rng.Height _
    )

    ' 画像データをコピー
    rng.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

    ' Chartに貼り付け
    chartObj.Chart.Paste

    ' PNG形式で出力
    chartObj.Chart.Export Filename:=outputPath, FilterName:="PNG"

    ' 一時オブジェクト削除
    chartObj.Delete

    ' ファイルを閉じる（開いた場合のみ）
    If needClose Then
        wb.Close SaveChanges:=False
    End If

    ' 成功ログ
    Call Module1_Main.LogMessage("  [OK] 画像生成完了: " & outputPath & " (" & Format(Timer - startTime, "0.00") & "秒)")
    ExportRangeToImage = True
    Exit Function

ErrorHandler:
    ' エラーハンドリング
    If Not chartObj Is Nothing Then chartObj.Delete
    If needClose And Not wb Is Nothing Then wb.Close SaveChanges:=False
    Call Module1_Main.LogMessage("  [ERROR] 画像生成失敗: " & Err.Description & " (Err#" & Err.Number & ")")
    ExportRangeToImage = False
End Function

'========================================
' パラメータ検証
'========================================
Private Function ValidateParameters( _
    sourceFilePath As String, _
    sourceSheetName As String, _
    sourceRange As String, _
    outputFolder As String, _
    outputFileName As String _
) As Boolean

    On Error GoTo ErrorHandler

    ' ファイル存在確認（外部ファイルの場合）
    If sourceFilePath <> "" Then
        If Dir(sourceFilePath) = "" Then
            Call Module1_Main.LogMessage("  [ERROR] ソースファイルが見つかりません: " & sourceFilePath)
            ValidateParameters = False
            Exit Function
        End If
    End If

    ' シート名検証
    If Trim(sourceSheetName) = "" Then
        Call Module1_Main.LogMessage("  [ERROR] シート名が空です")
        ValidateParameters = False
        Exit Function
    End If

    ' 範囲検証
    If Trim(sourceRange) = "" Then
        Call Module1_Main.LogMessage("  [ERROR] 範囲が空です")
        ValidateParameters = False
        Exit Function
    End If

    ' 出力フォルダ検証
    If Not FolderExists(outputFolder) Then
        Call Module1_Main.LogMessage("  [ERROR] 出力フォルダが見つかりません: " & outputFolder)
        ValidateParameters = False
        Exit Function
    End If

    ' ファイル名検証
    If Trim(outputFileName) = "" Then
        Call Module1_Main.LogMessage("  [ERROR] ファイル名が空です")
        ValidateParameters = False
        Exit Function
    End If

    ValidateParameters = True
    Exit Function

ErrorHandler:
    Call Module1_Main.LogMessage("  [ERROR] パラメータ検証エラー: " & Err.Description)
    ValidateParameters = False
End Function

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
' フォルダ存在確認
'========================================
Private Function FolderExists(folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> "")
    On Error GoTo 0
End Function

'========================================
' 簡易版ラッパー関数（テスト用）
'========================================
Public Function ExportRangeToImageSimple() As Boolean
    ' デフォルト値でテスト実行
    ExportRangeToImageSimple = ExportRangeToImage( _
        sourceFilePath:="", _
        sourceSheetName:="Sheet1", _
        sourceRange:="A1:D10", _
        outputFolder:="C:\Users\t-tsuji\AIアプリ開発\release-creator\files", _
        outputFileName:="test_chart.png" _
    )
End Function
