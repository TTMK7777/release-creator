Attribute VB_Name = "Module1_Main_Complete"
'========================================
' Module1_Main_Complete
' プレスリリース自動生成マクロ - メイン制御モジュール（完全版）
'
' 作成日: 2025年11月11日
' 作成者: Claude Code
' バージョン: 3.0 (Complete - 改善版Module3/4統合)
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
' メイン実行（完全版）
'========================================
Public Sub 実行完全版()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    Dim result As Boolean

    startTime = Timer
    ClearLog

    LogMessage "========================================="
    LogMessage "プレスリリース自動生成 開始 (v3.0 Complete)"
    LogMessage "========================================="

    ' ========================================
    ' Phase 1: データ転記
    ' ========================================
    LogMessage ""
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
    LogMessage "[ERROR] 予期しないエラー: " & Err.Description & " (Err#" & Err.Number & ")"
    Resume CleanupAndExit
End Sub

'========================================
' Phase 2: 画像生成
'========================================
Private Function GenerateImages() As Boolean
    On Error GoTo ErrorHandler

    Dim outputFolder As String
    Dim result As Boolean

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

    If Not result Then
        GenerateImages = False
        Exit Function
    End If

    ' 評価項目別ランキング表の画像化（オプション）
    ' 必要に応じてコメント解除
    ' LogMessage "  [INFO] 評価項目別ランキング表を画像化..."
    ' result = Module3_Image_Improved.ExportRangeToImage( _
    '     sourceSheetName:="評価項目別", _
    '     sourceRange:="B7:D20", _
    '     outputFolder:=outputFolder, _
    '     outputFileName:="評価項目別ランキング.png" _
    ' )

    GenerateImages = True
    Exit Function

ErrorHandler:
    LogMessage "  [ERROR] 画像生成エラー: " & Err.Description
    GenerateImages = False
End Function

'========================================
' Phase 3: Word文書更新
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

    UpdateWordDocuments = True
    Exit Function

ErrorHandler:
    LogMessage "  [ERROR] Word文書更新エラー: " & Err.Description
    UpdateWordDocuments = False
End Function

'========================================
' 簡易版実行（後方互換性のため）
'========================================
Public Sub 実行()
    ' Module2のみを実行する簡易版
    Dim startTime As Double
    startTime = Timer

    ClearLog

    LogMessage "========================================="
    LogMessage "プレスリリース自動生成 開始 (簡易版)"
    LogMessage "========================================="

    ' データ転記を実行
    Dim result As Boolean
    result = Module2_Data_Improved.TransferRankingDataSimple()

    If result Then
        LogMessage "[OK] 処理が正常に完了しました"
    Else
        LogMessage "[ERROR] 処理中にエラーが発生しました"
    End If

    Dim endTime As Double
    endTime = Timer

    LogMessage "実行時間: " & Format((endTime - startTime), "0.00") & "秒"
    LogMessage "========================================="

    ' ログを表示
    ShowLog
End Sub
