Attribute VB_Name = "Module1_Main_Optimized"
'========================================
' Module1_Main_Optimized
' プレスリリース自動生成マクロ - メイン制御モジュール（最適化版）
'
' 作成日: 2025年11月11日
' 作成者: Claude Code
' バージョン: 4.0 (Optimized - パフォーマンス最適化 + 一気通貫実行)
'========================================
Option Explicit

' グローバルログ変数
Private g_ExecutionLog As String
Private g_ErrorCount As Long
Private g_WarningCount As Long

'========================================
' ログ記録関数
'========================================
Public Sub LogMessage(msg As String)
    Dim timestamp As String
    timestamp = Format(Now, "hh:nn:ss")

    g_ExecutionLog = g_ExecutionLog & "[" & timestamp & "] " & msg & vbCrLf
    Debug.Print "[" & timestamp & "] " & msg

    ' エラーカウント
    If InStr(msg, "[ERROR]") > 0 Then
        g_ErrorCount = g_ErrorCount + 1
    ElseIf InStr(msg, "[WARN]") > 0 Then
        g_WarningCount = g_WarningCount + 1
    End If
End Sub

'========================================
' ログ表示
'========================================
Public Sub ShowLog()
    Dim summary As String
    summary = "エラー: " & g_ErrorCount & "件 | 警告: " & g_WarningCount & "件" & vbCrLf & vbCrLf
    MsgBox summary & g_ExecutionLog, vbInformation, "実行ログ"
End Sub

'========================================
' ログクリア
'========================================
Private Sub ClearLog()
    g_ExecutionLog = ""
    g_ErrorCount = 0
    g_WarningCount = 0
End Sub

'========================================
' メイン実行（最適化版 - 一気通貫）
'========================================
Public Sub 実行最適化版()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim result As Boolean

    ' パフォーマンス最適化用の変数
    Dim originalScreenUpdating As Boolean
    Dim originalCalculation As XlCalculation
    Dim originalEnableEvents As Boolean

    startTime = Timer
    ClearLog

    ' ========================================
    ' パフォーマンス設定の保存と最適化
    ' ========================================
    originalScreenUpdating = Application.ScreenUpdating
    originalCalculation = Application.Calculation
    originalEnableEvents = Application.EnableEvents

    ' 最適化設定を適用
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = True

    LogMessage "========================================="
    LogMessage "プレスリリース自動生成 開始 (v4.0 Optimized)"
    LogMessage "========================================="
    LogMessage "[INFO] パフォーマンス最適化モード: ON"
    LogMessage ""

    ' ========================================
    ' Phase 1: データ転記
    ' ========================================
    Application.StatusBar = "Phase 1/3: データ転記中..."
    LogMessage "[PHASE 1] データ転記を開始..."

    result = Module2_Data_Improved.TransferRankingDataSimple()

    If Not result Then
        LogMessage "  [ERROR] データ転記に失敗しました"
        GoTo CleanupAndExit
    End If

    LogMessage "  [OK] Phase 1 完了"

    ' ========================================
    ' Phase 2: 画像生成
    ' ========================================
    Application.StatusBar = "Phase 2/3: 画像生成中..."
    LogMessage ""
    LogMessage "[PHASE 2] 画像生成を開始..."

    result = GenerateImages()

    If Not result Then
        LogMessage "  [ERROR] 画像生成に失敗しました"
        GoTo CleanupAndExit
    End If

    LogMessage "  [OK] Phase 2 完了"

    ' ========================================
    ' Phase 3: Word文書更新
    ' ========================================
    Application.StatusBar = "Phase 3/3: Word文書更新中..."
    LogMessage ""
    LogMessage "[PHASE 3] Word文書更新を開始..."

    result = UpdateWordDocuments()

    If Not result Then
        LogMessage "  [ERROR] Word文書更新に失敗しました"
        GoTo CleanupAndExit
    End If

    LogMessage "  [OK] Phase 3 完了"

    ' ========================================
    ' 成功時の処理
    ' ========================================
CleanupAndExit:
    ' Application設定を必ず復元
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.StatusBar = False

    LogMessage ""
    LogMessage "========================================="

    If g_ErrorCount = 0 Then
        LogMessage "[OK] 全処理が正常に完了しました"
    Else
        LogMessage "[WARN] 処理中に " & g_ErrorCount & " 件のエラーが発生しました"
    End If

    Dim endTime As Double
    endTime = Timer
    LogMessage "実行時間: " & Format((endTime - startTime), "0.00") & "秒"
    LogMessage "========================================="

    ' ログを表示
    ShowLog
    Exit Sub

ErrorHandler:
    ' エラー時も必ず設定を復元
    Application.ScreenUpdating = originalScreenUpdating
    Application.Calculation = originalCalculation
    Application.EnableEvents = originalEnableEvents
    Application.StatusBar = False

    LogMessage "[ERROR] 予期しないエラー: " & Err.Description & " (Err#" & Err.Number & ")"
    Resume CleanupAndExit
End Sub

'========================================
' Phase 2: 画像生成（最適化版）
'========================================
Private Function GenerateImages() As Boolean
    On Error GoTo ErrorHandler

    Dim outputFolder As String
    Dim result As Boolean
    Dim imageCount As Long

    imageCount = 0

    ' 出力フォルダ設定
    outputFolder = ThisWorkbook.Path & "\files"

    ' フォルダ存在確認
    If Dir(outputFolder, vbDirectory) = "" Then
        LogMessage "  [INFO] 出力フォルダを作成: " & outputFolder
        MkDir outputFolder
    End If

    ' 総合ランキング表の画像化
    LogMessage "  [INFO] 総合ランキング表を画像化..."
    result = Module3_Image_Improved.ExportRangeToImage( _
        sourceSheetName:="総合3つ", _
        sourceRange:="B7:D15", _
        outputFolder:=outputFolder, _
        outputFileName:="総合ランキング.png" _
    )

    If result Then imageCount = imageCount + 1

    ' 評価項目別ランキング表の画像化（オプション）
    ' 必要に応じてコメント解除
    ' LogMessage "  [INFO] 評価項目別ランキング表を画像化..."
    ' result = Module3_Image_Improved.ExportRangeToImage( _
    '     sourceSheetName:="評価項目別", _
    '     sourceRange:="B7:D20", _
    '     outputFolder:=outputFolder, _
    '     outputFileName:="評価項目別ランキング.png" _
    ' )
    '
    ' If result Then imageCount = imageCount + 1

    LogMessage "  [INFO] 画像生成完了: " & imageCount & " 件"
    GenerateImages = (imageCount > 0)
    Exit Function

ErrorHandler:
    LogMessage "  [ERROR] 画像生成エラー: " & Err.Description
    GenerateImages = False
End Function

'========================================
' Phase 3: Word文書更新（最適化版）
'========================================
Private Function UpdateWordDocuments() As Boolean
    On Error GoTo ErrorHandler

    Dim templatePath As String
    Dim savePath As String
    Dim imageFilePath As String
    Dim result As Boolean

    ' パス設定
    templatePath = ThisWorkbook.Path & "\テンプレート\【テンプレ】20XX年X月発表 オリコン顧客満足度(R)調査 ●● ニュースリリース.docx"
    savePath = ThisWorkbook.Path & "\files\2025年携帯キャリアランキング_ニュースリリース.docx"
    imageFilePath = ThisWorkbook.Path & "\files\総合ランキング.png"

    ' Word文書更新
    LogMessage "  [INFO] Word文書を更新中..."
    result = Module4_Word_Improved.UpdateWordDocument( _
        templatePath:=templatePath, _
        savePath:=savePath, _
        rankingYear:="2025", _
        rankingName:="携帯キャリア", _
        totalRespondents:=8464, _
        imageBookmark:="RankingChart", _
        imageFilePath:=imageFilePath _
    )

    If Not result Then
        UpdateWordDocuments = False
        Exit Function
    End If

    LogMessage "  [INFO] 文書保存完了: " & savePath

    UpdateWordDocuments = True
    Exit Function

ErrorHandler:
    LogMessage "  [ERROR] Word文書更新エラー: " & Err.Description
    UpdateWordDocuments = False
End Function

'========================================
' ファイルベースのログ出力（オプション）
'========================================
Public Sub LogToFile(ByVal logLevel As String, ByVal message As String)
    On Error Resume Next

    Dim logFile As String
    Dim fileNum As Integer

    logFile = ThisWorkbook.Path & "\logs\ReleaseCreator_" & _
              Format(Date, "yyyymmdd") & ".log"

    ' ログフォルダ作成
    If Dir(ThisWorkbook.Path & "\logs", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\logs"
    End If

    fileNum = FreeFile
    Open logFile For Append As #fileNum

    Print #fileNum, Format(Now, "yyyy-mm-dd hh:mm:ss") & _
                    " [" & logLevel & "] " & message

    Close #fileNum
End Sub

'========================================
' 簡易版実行（後方互換性のため）
'========================================
Public Sub 実行()
    ' 最適化版を呼び出し
    Call 実行最適化版
End Sub

'========================================
' バッチ処理用（複数ランキング対応）
'========================================
Public Sub 実行バッチ処理(ParamArray rankings() As Variant)
    ' 将来の拡張用
    ' 複数のランキング（携帯キャリア、ネット銀行など）を一括処理
    ' 例: Call 実行バッチ処理("携帯キャリア", "ネット銀行", "クレジットカード")
End Sub
