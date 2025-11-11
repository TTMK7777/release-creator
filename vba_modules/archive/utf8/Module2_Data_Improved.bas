Attribute VB_Name = "Module2_Data"
'========================================
' Module2_Data: データ転記モジュール (改善版)
' ベース: マクロ記録コード (2025/11/10)
' 改善: 動的対応、エラーハンドリング、ログ機能
'========================================

Option Explicit

'========================================
' メインデータ転記関数
'========================================
Public Function TransferRankingData( _
    sourceFilePath As String, _
    targetFilePath As String, _
    rankingYear As String, _
    rankingName As String, _
    totalRespondents As Long, _
    rankingCount As Integer _
) As Boolean

    On Error GoTo ErrorHandler

    TransferRankingData = False

    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet

    Module1_Main.LogMessage "データ転記を開始します..."
    Module1_Main.LogMessage "  元データ: " & sourceFilePath
    Module1_Main.LogMessage "  出力先: " & targetFilePath

    ' ファイルを開く
    Module1_Main.LogMessage "  [DEBUG] ソースファイルを開いています..."
    Set sourceWb = Workbooks.Open(sourceFilePath, ReadOnly:=True)
    Module1_Main.LogMessage "  [DEBUG] ソースファイルを開きました"

    Module1_Main.LogMessage "  [DEBUG] ターゲットファイルを開いています..."
    Set targetWb = Workbooks.Open(targetFilePath)
    Module1_Main.LogMessage "  [DEBUG] ターゲットファイルを開きました"

    ' シートを取得 (記録から判明した実際のシート名)
    Module1_Main.LogMessage "  [DEBUG] シートを取得しています..."
    Set sourceWs = sourceWb.Worksheets("総合対象企業")  ' または適切なシート名
    Module1_Main.LogMessage "  [DEBUG] ソースシート取得: 総合対象企業"

    Set targetWs = targetWb.Worksheets("総合3つ")        ' または適切なシート名
    Module1_Main.LogMessage "  [DEBUG] ターゲットシート取得: 総合3つ"

    Module1_Main.LogMessage "  ランキング企業数: " & rankingCount & "社"

    ' ================================================
    ' 1. ランキングデータの転記 (動的に対応)
    ' ================================================
    Dim i As Long
    Dim sourceRow As Long
    Dim targetRow As Long

    sourceRow = 5  ' 記録から判明: 元データは5行目から開始
    targetRow = 11 ' 記録から判明: リリース表は11行目から開始

    For i = 1 To rankingCount
        ' 企業名を転記 (C列 → C列)
        targetWs.Cells(targetRow, 3).Value = sourceWs.Cells(sourceRow, 3).Value

        ' 得点を転記 (G列 → D列) - 値のみ貼り付け
        targetWs.Cells(targetRow, 4).Value = sourceWs.Cells(sourceRow, 7).Value

        sourceRow = sourceRow + 1
        targetRow = targetRow + 1
    Next i

    Module1_Main.LogMessage "  ✓ ランキングデータ転記完了 (" & rankingCount & "件)"

    ' ================================================
    ' 2. タイトルの動的生成 (記録から判明: C7:D9)
    ' ================================================
    Module1_Main.LogMessage "  [DEBUG] タイトル生成開始..."
    Dim titleText As String
    titleText = rankingYear & "年 オリコン顧客満足度®調査" & vbLf & _
                rankingName & " 総合ランキング (回答者数：" & _
                Format(totalRespondents, "#,##0") & "名)"
    Module1_Main.LogMessage "  [DEBUG] タイトルテキスト生成完了 (長さ: " & Len(titleText) & "文字)"

    ' タイトルセルに設定
    Module1_Main.LogMessage "  [DEBUG] タイトルをセルに設定中..."
    targetWs.Range("C7:D9").Value = titleText
    Module1_Main.LogMessage "  [DEBUG] セルへの設定完了"

    ' フォント書式設定 (記録されたフォーマットを適用)
    Module1_Main.LogMessage "  [DEBUG] フォント書式設定中..."
    With targetWs.Range("C7:D9")
        .Font.Name = "BIZ UDPゴシック"
        .Font.Bold = True
        .Font.Size = 12
    End With
    Module1_Main.LogMessage "  [DEBUG] フォント書式設定完了"

    ' ®マークを上付き文字に (動的に検索)
    Module1_Main.LogMessage "  [DEBUG] ®マーク検索中..."
    Dim regMarkPos As Long
    regMarkPos = InStr(titleText, "®")
    Module1_Main.LogMessage "  [DEBUG] ®マーク位置: " & regMarkPos
    If regMarkPos > 0 Then
        Module1_Main.LogMessage "  [DEBUG] ®マークを上付き文字に設定中..."
        targetWs.Range("C7:D9").Characters(regMarkPos, 1).Font.Superscript = True
        Module1_Main.LogMessage "  [DEBUG] 上付き文字設定完了"
    End If

    Module1_Main.LogMessage "  ✓ タイトル生成完了"

    ' ================================================
    ' 3. 注釈の動的生成 (記録から判明: B15:D15)
    ' ================================================
    Dim currentYear As Integer
    Dim currentMonth As Integer
    Dim publishDate As String

    currentYear = Year(Now)
    currentMonth = Month(Now)
    publishDate = currentYear & "年" & currentMonth & "月1日"

    Dim noteText As String
    noteText = "※上記順位以降はサイトに掲載しております。" & vbLf & _
               "調査主体：株式会社oricon ME（" & publishDate & "発表）"

    targetWs.Range("B15:D15").Value = noteText

    Module1_Main.LogMessage "  ✓ 注釈生成完了"

    ' ================================================
    ' 4. 保存とクリーンアップ
    ' ================================================
    targetWb.Save

    targetWb.Close SaveChanges:=True
    sourceWb.Close SaveChanges:=False

    Set targetWs = Nothing
    Set sourceWs = Nothing
    Set targetWb = Nothing
    Set sourceWb = Nothing

    Module1_Main.LogMessage "データ転記が正常に完了しました"
    TransferRankingData = True

    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "[ERROR] データ転記エラー: " & Err.Description
    Module1_Main.LogMessage "        エラー番号: " & Err.Number
    Module1_Main.LogMessage "        エラー発生行: " & Erl
    Module1_Main.LogMessage "        エラー発生元: " & Err.Source

    ' クリーンアップ
    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    If Not targetWb Is Nothing Then targetWb.Close SaveChanges:=False
    On Error GoTo 0

    TransferRankingData = False
End Function

'========================================
' 簡易呼び出し用ラッパー関数
'========================================
Public Function TransferRankingDataSimple() As Boolean
    ' デフォルト値で実行
    TransferRankingDataSimple = TransferRankingData( _
        sourceFilePath:="C:\Users\t-tsuji\AIアプリ開発\release-creator\テンプレート\【資料】携帯キャリア_ランキング結果2024.xlsx", _
        targetFilePath:="C:\Users\t-tsuji\AIアプリ開発\release-creator\テンプレート\【テンプレ】リリース内表.xlsx", _
        rankingYear:="2025", _
        rankingName:="携帯キャリア", _
        totalRespondents:=8464, _
        rankingCount:=4 _
    )
End Function

'========================================
' データ検証関数
'========================================
Public Function ValidateSourceData(sourceFilePath As String) As Boolean
    On Error GoTo ErrorHandler

    ValidateSourceData = False

    Dim wb As Workbook
    Dim ws As Worksheet

    Module1_Main.LogMessage "元データの検証を開始します..."

    Set wb = Workbooks.Open(sourceFilePath, ReadOnly:=True)
    Set ws = wb.Worksheets("総合対象企業")

    ' データ存在チェック (C5セルに企業名があるか)
    If IsEmpty(ws.Cells(5, 3).Value) Or ws.Cells(5, 3).Value = "" Then
        Module1_Main.LogMessage "エラー: 1位の企業名が見つかりません (C5)"
        wb.Close SaveChanges:=False
        Exit Function
    End If

    ' 得点チェック (G5セルに数値があるか)
    If Not IsNumeric(ws.Cells(5, 7).Value) Then
        Module1_Main.LogMessage "エラー: 1位の得点が数値ではありません (G5)"
        wb.Close SaveChanges:=False
        Exit Function
    End If

    wb.Close SaveChanges:=False

    Module1_Main.LogMessage "✓ 元データの検証が完了しました"
    ValidateSourceData = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "データ検証エラー: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ValidateSourceData = False
End Function
