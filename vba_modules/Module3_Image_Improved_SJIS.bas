Attribute VB_Name = "Module3_Image_Improved"
'========================================
' Module3_Image_Improved
' Excel範囲を画像化するモジュール（改善版）
'
' 作成日: 2025-11-11
' 作成者: GPT-4o提案 + Claude Code実装
' バージョン: 2.0 (Improved)
'========================================
Option Explicit

'========================================
' 公開関数: 範囲を画像として出力
'========================================
' @param sourceSheetName  ソースシート名
' @param sourceRange      ソース範囲（例: "B7:D15"）
' @param outputFolder     出力フォルダパス
' @param outputFileName   出力ファイル名（例: "RankingChart.png"）
' @return Boolean         成功時True、失敗時False
'========================================
Public Function ExportRangeToImage( _
    sourceSheetName As String, _
    sourceRange As String, _
    outputFolder As String, _
    outputFileName As String _
) As Boolean

    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rng As Range
    Dim chartObj As ChartObject
    Dim outputPath As String
    Dim startTime As Double

    startTime = Timer
    Call Module1_Main.LogMessage("画像生成を開始: " & sourceSheetName & "!" & sourceRange)

    ' パラメータ検証
    If Not ValidateParameters(sourceSheetName, sourceRange, outputFolder, outputFileName) Then
        ExportRangeToImage = False
        Exit Function
    End If

    ' シート取得
    Set ws = ThisWorkbook.Worksheets(sourceSheetName)

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

    ' 成功ログ
    Call Module1_Main.LogMessage("  [OK] 画像生成完了: " & outputPath & " (" & Format(Timer - startTime, "0.00") & "秒)")
    ExportRangeToImage = True
    Exit Function

ErrorHandler:
    ' エラーハンドリング
    If Not chartObj Is Nothing Then chartObj.Delete
    Call Module1_Main.LogMessage("  [ERROR] 画像生成失敗: " & Err.Description & " (Err#" & Err.Number & ")")
    ExportRangeToImage = False
End Function

'========================================
' パラメータ検証
'========================================
Private Function ValidateParameters( _
    sourceSheetName As String, _
    sourceRange As String, _
    outputFolder As String, _
    outputFileName As String _
) As Boolean

    On Error GoTo ErrorHandler

    ' シート存在確認
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sourceSheetName)
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Call Module1_Main.LogMessage("  [ERROR] シートが見つかりません: " & sourceSheetName)
        ValidateParameters = False
        Exit Function
    End If

    ' 範囲検証
    Dim testRange As Range
    On Error Resume Next
    Set testRange = ws.Range(sourceRange)
    On Error GoTo ErrorHandler

    If testRange Is Nothing Then
        Call Module1_Main.LogMessage("  [ERROR] 無効な範囲: " & sourceRange)
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
        sourceSheetName:="Sheet1", _
        sourceRange:="A1:D10", _
        outputFolder:="C:\Users\t-tsuji\AIアプリ開発\release-creator\files", _
        outputFileName:="test_chart.png" _
    )
End Function
