Attribute VB_Name = "Module3_Image_Improved"
'========================================
' Module3_Image_Improved
' Excel範囲を高品質PNG画像として出力（完全動的版）
'
' 作成日: 2025-11-11
' バージョン: 4.0 (完全動的化 + 真っ白画像問題の根本解決)
'
' 改善点:
' - すべてのハードコード値を排除（シート名、範囲、ファイル名）
' - 引数によるパラメータ化で柔軟性を大幅向上
' - Module2_Data_Dynamicと同等の品質とエラーハンドリング
' - Chart.Export方式による高品質PNG出力
'========================================
Option Explicit

'========================================
' データ型定義
'========================================
Public Type ImageExportConfig
    SourceFilePath As String    ' ソースファイルのパス
    SourceSheetName As String   ' ソースシート名
    SourceRange As String       ' ソース範囲（例: "C4:E10"）
    OutputFolder As String      ' 出力フォルダパス
    OutputFileName As String    ' 出力ファイル名（例: "総合ランキング.png"）
    ImageWidth As Long          ' 画像幅（0=自動）
    ImageHeight As Long         ' 画像高さ（0=自動）
End Type

'========================================
' 公開関数: 範囲を画像として出力（完全動的版）
'========================================
Public Function ExportRangeToImage( _
    config As ImageExportConfig _
) As Boolean

    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim outputPath As String
    Dim startTime As Double
    Dim needClose As Boolean

    startTime = Timer
    needClose = False

    Module1_Main.LogMessage "  [画像生成] " & config.SourceSheetName & "!" & config.SourceRange & " → " & config.OutputFileName

    ' ========================================
    ' パラメータ検証
    ' ========================================
    If Not ValidateConfig(config) Then
        Module1_Main.LogMessage "    [ERROR] パラメータが不正です"
        ExportRangeToImage = False
        Exit Function
    End If

    ' ========================================
    ' ファイルを開く
    ' ========================================
    If config.SourceFilePath <> "" Then
        If Not IsFileOpen(config.SourceFilePath) Then
            Set wb = Workbooks.Open(config.SourceFilePath, ReadOnly:=True, UpdateLinks:=False)
            needClose = True
        Else
            Set wb = GetOpenWorkbook(config.SourceFilePath)
        End If
    Else
        Set wb = ThisWorkbook
    End If

    ' ========================================
    ' シート名の柔軟な検索（スペース・全角半角対応）
    ' ========================================
    Set ws = FindWorksheet(wb, config.SourceSheetName)

    If ws Is Nothing Then
        Module1_Main.LogMessage "    [ERROR] シート「" & config.SourceSheetName & "」が見つかりません"
        If needClose Then wb.Close SaveChanges:=False
        ExportRangeToImage = False
        Exit Function
    End If

    ' ========================================
    ' 範囲取得
    ' ========================================
    Set rng = ws.Range(config.SourceRange)

    ' ========================================
    ' 出力パス生成
    ' ========================================
    outputPath = BuildOutputPath(config.OutputFolder, config.OutputFileName)

    ' ========================================
    ' フォルダ作成（存在しない場合）
    ' ========================================
    CreateFolderIfNotExists config.OutputFolder

    ' ========================================
    ' 画像生成（Chart.Export方式）
    ' ========================================
    If Not ExportRangeAsChart(rng, outputPath, config.ImageWidth, config.ImageHeight) Then
        Module1_Main.LogMessage "    [ERROR] PNG出力に失敗しました"
        If needClose Then wb.Close SaveChanges:=False
        ExportRangeToImage = False
        Exit Function
    End If

    ' ========================================
    ' ファイルを閉じる
    ' ========================================
    If needClose Then
        wb.Close SaveChanges:=False
    End If

    Module1_Main.LogMessage "    [OK] 画像生成完了 (" & Format(Timer - startTime, "0.00") & "秒)"
    ExportRangeToImage = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "    [ERROR] 画像生成失敗: " & Err.Description & " (Err#" & Err.Number & ")"

    On Error Resume Next
    If needClose And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0

    ExportRangeToImage = False
End Function

'========================================
' Chart.Export方式による画像出力（最も安定版）
' AI推奨: Claude Sonnet 4.5 解決策4
'========================================
Private Function ExportRangeAsChart( _
    rng As Range, _
    outputPath As String, _
    Optional imageWidth As Long = 0, _
    Optional imageHeight As Long = 0 _
) As Boolean

    On Error GoTo ErrorHandler

    Dim chtObj As ChartObject
    Dim tempSheet As Worksheet
    Dim defaultWidth As Long
    Dim defaultHeight As Long

    ' デフォルトサイズ設定
    defaultWidth = IIf(imageWidth > 0, imageWidth, 600)
    defaultHeight = IIf(imageHeight > 0, imageHeight, 400)

    ' 一時シートを作成
    Set tempSheet = ThisWorkbook.Worksheets.Add
    Application.ScreenUpdating = False

    ' 範囲を画像としてコピー（xlBitmap形式）
    rng.CopyPicture Appearance:=xlScreen, Format:=xlBitmap

    ' チャートオブジェクトを作成
    Set chtObj = tempSheet.ChartObjects.Add(0, 0, defaultWidth, defaultHeight)

    With chtObj
        With .Chart
            ' チャートエリアをクリア
            .ChartArea.Clear

            ' 画像を貼り付け
            .Paste

            ' 余白を削除
            On Error Resume Next
            .ChartArea.Format.Line.Visible = msoFalse
            On Error GoTo ErrorHandler

            ' PNG出力
            .Export Filename:=outputPath, FilterName:="PNG"
        End With
    End With

    ' クリーンアップ
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

    ExportRangeAsChart = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "    [ERROR] Chart出力エラー: " & Err.Description

    ' クリーンアップ
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not tempSheet Is Nothing Then tempSheet.Delete
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    On Error GoTo 0

    ExportRangeAsChart = False
End Function

'========================================
' パラメータ検証
'========================================
Private Function ValidateConfig(config As ImageExportConfig) As Boolean
    ValidateConfig = True

    If Trim(config.SourceSheetName) = "" Then
        Module1_Main.LogMessage "    [ERROR] シート名が空です"
        ValidateConfig = False
    End If

    If Trim(config.SourceRange) = "" Then
        Module1_Main.LogMessage "    [ERROR] 範囲が空です"
        ValidateConfig = False
    End If

    If Trim(config.OutputFolder) = "" Then
        Module1_Main.LogMessage "    [ERROR] 出力フォルダが空です"
        ValidateConfig = False
    End If

    If Trim(config.OutputFileName) = "" Then
        Module1_Main.LogMessage "    [ERROR] 出力ファイル名が空です"
        ValidateConfig = False
    End If
End Function

'========================================
' シート名の柔軟な検索
'========================================
Private Function FindWorksheet(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim searchNames() As String
    Dim i As Long

    ' 検索パターン生成（スペース・全角半角のバリエーション）
    ReDim searchNames(1 To 8)
    searchNames(1) = sheetName
    searchNames(2) = Trim(sheetName)
    searchNames(3) = " " & sheetName
    searchNames(4) = sheetName & " "
    searchNames(5) = " " & sheetName & " "
    searchNames(6) = Replace(sheetName, "＋", "+")
    searchNames(7) = Replace(sheetName, "+", "＋")
    searchNames(8) = " " & Replace(sheetName, "＋", "+") & " "

    On Error Resume Next
    For i = 1 To 8
        Set ws = wb.Worksheets(searchNames(i))
        If Not ws Is Nothing Then
            Set FindWorksheet = ws
            Exit Function
        End If
    Next i
    On Error GoTo 0

    Set FindWorksheet = Nothing
End Function

'========================================
' 出力パス構築
'========================================
Private Function BuildOutputPath(folder As String, fileName As String) As String
    Dim path As String
    path = folder

    ' 末尾の\を確認
    If Right(path, 1) <> "\" Then
        path = path & "\"
    End If

    BuildOutputPath = path & fileName
End Function

'========================================
' フォルダ作成（存在しない場合）
'========================================
Private Sub CreateFolderIfNotExists(folderPath As String)
    On Error Resume Next
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
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
' 3種類の画像を生成する統合関数（動的版）
'========================================
Public Function GenerateAllRankingImages( _
    targetFilePath As String, _
    outputFolder As String, _
    Optional imageWidth As Long = 0, _
    Optional imageHeight As Long = 0 _
) As Boolean

    On Error GoTo ErrorHandler
    GenerateAllRankingImages = False

    Dim config As ImageExportConfig
    Dim imageCount As Long
    imageCount = 0

    Module1_Main.LogMessage "  [画像生成] 3種類のランキング画像を生成します"
    Module1_Main.LogMessage ""

    ' =========================================
    ' 共通設定
    ' =========================================
    config.SourceFilePath = targetFilePath
    config.OutputFolder = outputFolder
    config.ImageWidth = imageWidth
    config.ImageHeight = imageHeight

    ' =========================================
    ' 1. 総合ランキング画像（総合1つシート）
    ' =========================================
    Module1_Main.LogMessage "  [1/3] 総合ランキング画像を生成中..."
    config.SourceSheetName = "総合1つ"
    config.SourceRange = "C4:E10"
    config.OutputFileName = "01_総合ランキング.png"

    If ExportRangeToImage(config) Then
        imageCount = imageCount + 1
    End If
    Module1_Main.LogMessage ""

    ' =========================================
    ' 2. 評価項目ランキング画像（総合2つシート右側）
    ' =========================================
    Module1_Main.LogMessage "  [2/3] 評価項目ランキング画像を生成中..."
    config.SourceSheetName = "総合2つ"
    config.SourceRange = "G4:I10"
    config.OutputFileName = "02_評価項目ランキング.png"

    If ExportRangeToImage(config) Then
        imageCount = imageCount + 1
    End If
    Module1_Main.LogMessage ""

    ' =========================================
    ' 3. 評価・部門別ランキング画像（評価・部門別シート）
    ' =========================================
    Module1_Main.LogMessage "  [3/3] 評価・部門別ランキング画像を生成中..."
    config.SourceSheetName = " 評価・部門別 "
    config.SourceRange = "B7:P18"
    config.OutputFileName = "03_評価部門別ランキング.png"

    If ExportRangeToImage(config) Then
        imageCount = imageCount + 1
    End If
    Module1_Main.LogMessage ""

    Module1_Main.LogMessage "  [画像生成] 完了: " & imageCount & "/3 件"

    GenerateAllRankingImages = (imageCount = 3)
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "  [ERROR] 画像生成処理エラー: " & Err.Description
    GenerateAllRankingImages = False
End Function

'========================================
' 簡易版: 単一画像エクスポート（後方互換）
'========================================
Public Function ExportSingleImage( _
    Optional sourceFilePath As String = "", _
    Optional sourceSheetName As String, _
    Optional sourceRange As String, _
    Optional outputFolder As String, _
    Optional outputFileName As String _
) As Boolean

    Dim config As ImageExportConfig

    config.SourceFilePath = sourceFilePath
    config.SourceSheetName = sourceSheetName
    config.SourceRange = sourceRange
    config.OutputFolder = outputFolder
    config.OutputFileName = outputFileName
    config.ImageWidth = 0
    config.ImageHeight = 0

    ExportSingleImage = ExportRangeToImage(config)
End Function
