Attribute VB_Name = "Module4_Word_Improved"
'========================================
' Module4_Word_Improved
' Wordドキュメント操作モジュール（改善版）
'
' 作成日: 2025-11-11
' 作成者: GPT-4o提案 + Claude Code実装
' バージョン: 2.0 (Improved)
'========================================
Option Explicit

'========================================
' 公開関数: Wordドキュメント更新
'========================================
' @param templatePath      テンプレートファイルパス
' @param savePath          保存先ファイルパス
' @param rankingYear       ランキング年度（例: "2025"）
' @param rankingName       ランキング名（例: "携帯キャリア"）
' @param totalRespondents  総回答者数
' @param imageBookmark     画像ブックマーク名（例: "RankingChart"）
' @param imageFilePath     画像ファイルパス
' @return Boolean          成功時True、失敗時False
'========================================
Public Function UpdateWordDocument( _
    templatePath As String, _
    savePath As String, _
    rankingYear As String, _
    rankingName As String, _
    totalRespondents As Long, _
    imageBookmark As String, _
    imageFilePath As String _
) As Boolean

    On Error GoTo ErrorHandler

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim startTime As Double

    startTime = Timer
    Call Module1_Main.LogMessage("Word文書更新を開始: " & templatePath)

    ' パラメータ検証
    If Not ValidateParameters(templatePath, savePath, imageFilePath) Then
        UpdateWordDocument = False
        Exit Function
    End If

    ' Word起動
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    ' テンプレート開く
    Set wdDoc = wdApp.Documents.Open(templatePath)

    ' 年度置換
    Call ReplaceText(wdDoc, "{{YEAR}}", rankingYear)

    ' ランキング名置換
    Call ReplaceText(wdDoc, "{{RANKING_NAME}}", rankingName)

    ' 回答者数置換
    Call ReplaceText(wdDoc, "{{TOTAL_RESPONDENTS}}", Format(totalRespondents, "#,##0"))

    ' 発表日置換
    Call ReplaceText(wdDoc, "{{RELEASE_DATE}}", Format(Date, "yyyy年m月d日"))

    ' 画像置換（ブックマーク使用）
    If imageBookmark <> "" And imageFilePath <> "" Then
        Call ReplaceImageAtBookmark(wdDoc, imageBookmark, imageFilePath)
    End If

    ' 保存
    wdDoc.SaveAs2 Filename:=savePath
    wdDoc.Close
    wdApp.Quit

    ' 成功ログ
    Call Module1_Main.LogMessage("  [OK] Word文書更新完了: " & savePath & " (" & Format(Timer - startTime, "0.00") & "秒)")
    UpdateWordDocument = True
    Exit Function

ErrorHandler:
    ' クリーンアップ
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit

    Call Module1_Main.LogMessage("  [ERROR] Word文書更新失敗: " & Err.Description & " (Err#" & Err.Number & ")")
    UpdateWordDocument = False
End Function

'========================================
' テキスト置換
'========================================
Private Sub ReplaceText( _
    wdDoc As Object, _
    findText As String, _
    replaceText As String _
)
    On Error GoTo ErrorHandler

    With wdDoc.Content.Find
        .Text = findText
        .Replacement.Text = replaceText
        .Execute Replace:=2  ' wdReplaceAll
    End With

    Exit Sub

ErrorHandler:
    Call Module1_Main.LogMessage("  [WARN] テキスト置換エラー: " & findText & " -> " & Err.Description)
End Sub

'========================================
' ブックマーク位置に画像を挿入
'========================================
Private Sub ReplaceImageAtBookmark( _
    wdDoc As Object, _
    bookmarkName As String, _
    imageFilePath As String _
)
    On Error GoTo ErrorHandler

    Dim bmkRange As Object

    ' ブックマーク存在確認
    If Not wdDoc.Bookmarks.Exists(bookmarkName) Then
        Call Module1_Main.LogMessage("  [WARN] ブックマークが見つかりません: " & bookmarkName)
        Exit Sub
    End If

    ' ブックマーク範囲取得
    Set bmkRange = wdDoc.Bookmarks(bookmarkName).Range

    ' 既存の画像を削除
    If bmkRange.InlineShapes.Count > 0 Then
        bmkRange.InlineShapes(1).Delete
    End If

    ' 画像挿入
    bmkRange.InlineShapes.AddPicture _
        Filename:=imageFilePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True

    Call Module1_Main.LogMessage("  [OK] 画像置換完了: " & bookmarkName)
    Exit Sub

ErrorHandler:
    Call Module1_Main.LogMessage("  [ERROR] 画像置換エラー: " & bookmarkName & " -> " & Err.Description)
End Sub

'========================================
' パラメータ検証
'========================================
Private Function ValidateParameters( _
    templatePath As String, _
    savePath As String, _
    imageFilePath As String _
) As Boolean

    On Error GoTo ErrorHandler

    ' テンプレートファイル存在確認
    If Not FileExists(templatePath) Then
        Call Module1_Main.LogMessage("  [ERROR] テンプレートファイルが見つかりません: " & templatePath)
        ValidateParameters = False
        Exit Function
    End If

    ' 保存先フォルダ存在確認
    Dim saveFolder As String
    saveFolder = Left(savePath, InStrRev(savePath, "\"))
    If Not FolderExists(saveFolder) Then
        Call Module1_Main.LogMessage("  [ERROR] 保存先フォルダが見つかりません: " & saveFolder)
        ValidateParameters = False
        Exit Function
    End If

    ' 画像ファイル存在確認（指定がある場合のみ）
    If imageFilePath <> "" Then
        If Not FileExists(imageFilePath) Then
            Call Module1_Main.LogMessage("  [ERROR] 画像ファイルが見つかりません: " & imageFilePath)
            ValidateParameters = False
            Exit Function
        End If
    End If

    ValidateParameters = True
    Exit Function

ErrorHandler:
    Call Module1_Main.LogMessage("  [ERROR] パラメータ検証エラー: " & Err.Description)
    ValidateParameters = False
End Function

'========================================
' ファイル存在確認
'========================================
Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
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
Public Function UpdateWordDocumentSimple() As Boolean
    ' デフォルト値でテスト実行
    UpdateWordDocumentSimple = UpdateWordDocument( _
        templatePath:="C:\Users\t-tsuji\AIアプリ開発\release-creator\テンプレート\【テンプレ】20XX年X月発表 オリコン顧客満足度(R)調査 ●● ニュースリリース.docx", _
        savePath:="C:\Users\t-tsuji\AIアプリ開発\release-creator\files\test_output.docx", _
        rankingYear:="2025", _
        rankingName:="携帯キャリア", _
        totalRespondents:=8464, _
        imageBookmark:="RankingChart", _
        imageFilePath:="C:\Users\t-tsuji\AIアプリ開発\release-creator\files\test_chart.png" _
    )
End Function
