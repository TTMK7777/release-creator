Attribute VB_Name = "Module2_Data"
'========================================
' Module2_Data_Complete: 完全な転記モジュール
' 作成日: 2025年11月11日
' バージョン: 5.0 (完全版)
'========================================
Option Explicit

'========================================
' メインデータ転記関数（完全版）
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

    Module1_Main.LogMessage "======================================"
    Module1_Main.LogMessage "完全データ転記を開始"
    Module1_Main.LogMessage "======================================"
    Module1_Main.LogMessage "  元データ: " & sourceFilePath
    Module1_Main.LogMessage "  転記先: " & targetFilePath
    Module1_Main.LogMessage ""

    ' ファイルを開く
    Set sourceWb = Workbooks.Open(sourceFilePath, ReadOnly:=True)
    Set targetWb = Workbooks.Open(targetFilePath)

    ' ================================================
    ' Phase 1: 総合ランキング転記 (3シート)
    ' ================================================
    Module1_Main.LogMessage "[Phase 1] 総合ランキング転記..."

    If Not Transfer_Overall_1Column(sourceWb, targetWb, rankingYear, rankingName, totalRespondents) Then
        GoTo ErrorHandler
    End If

    If Not Transfer_Overall_2Columns(sourceWb, targetWb, rankingYear, rankingName, totalRespondents) Then
        GoTo ErrorHandler
    End If

    If Not Transfer_Overall_WithPrevRank(sourceWb, targetWb, rankingYear, rankingName, totalRespondents) Then
        GoTo ErrorHandler
    End If

    Module1_Main.LogMessage "  [OK] Phase 1 完了"
    Module1_Main.LogMessage ""

    ' ================================================
    ' Phase 2: 評価項目別ランキング転記
    ' ================================================
    Module1_Main.LogMessage "[Phase 2] 評価項目別ランキング転記..."

    If Not Transfer_EvaluationItems(sourceWb, targetWb, rankingYear, rankingName) Then
        GoTo ErrorHandler
    End If

    Module1_Main.LogMessage "  [OK] Phase 2 完了"
    Module1_Main.LogMessage ""

    ' ================================================
    ' 保存とクリーンアップ
    ' ================================================
    targetWb.Save
    targetWb.Close SaveChanges:=True
    sourceWb.Close SaveChanges:=False

    Set targetWb = Nothing
    Set sourceWb = Nothing

    Module1_Main.LogMessage "======================================"
    Module1_Main.LogMessage "全転記処理が正常に完了しました"
    Module1_Main.LogMessage "======================================"
    TransferRankingData = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "[ERROR] 転記エラー: " & Err.Description
    Module1_Main.LogMessage "        エラー番号: " & Err.Number

    On Error Resume Next
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    If Not targetWb Is Nothing Then targetWb.Close SaveChanges:=False
    On Error GoTo 0

    TransferRankingData = False
End Function

'========================================
' 総合ランキング転記: 総合1つシート
'========================================
Private Function Transfer_Overall_1Column( _
    sourceWb As Workbook, _
    targetWb As Workbook, _
    rankingYear As String, _
    rankingName As String, _
    totalRespondents As Long _
) As Boolean

    On Error GoTo ErrorHandler
    Transfer_Overall_1Column = False

    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim i As Long

    Set sourceWs = sourceWb.Worksheets("総合対象企業")
    Set targetWs = targetWb.Worksheets("総合1つ")

    Module1_Main.LogMessage "  [1/3] 総合1つシートに転記中..."

    ' タイトル設定 (C4セル)
    Dim titleText As String
    titleText = rankingName & " (回答者数：" & Format(totalRespondents, "#,##0") & "名)"
    targetWs.Range("C4").Value = titleText

    ' データ転記 (D6:E10 - 最大5位まで)
    Dim maxRows As Long
    maxRows = Application.Min(5, sourceWs.Cells(sourceWs.Rows.Count, 3).End(xlUp).Row - 4)

    For i = 1 To maxRows
        ' 企業名 (C列 → D列)
        targetWs.Cells(5 + i, 4).Value = sourceWs.Cells(4 + i, 3).Value

        ' 得点 (G列 → E列) - 小数点以下2桁
        targetWs.Cells(5 + i, 5).Value = Round(sourceWs.Cells(4 + i, 7).Value, 2)
    Next i

    Module1_Main.LogMessage "    [OK] 総合1つシート完了 (" & maxRows & "社転記)"
    Transfer_Overall_1Column = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "    [ERROR] 総合1つシート転記エラー: " & Err.Description
    Transfer_Overall_1Column = False
End Function

'========================================
' 総合ランキング転記: 総合2つシート
'========================================
Private Function Transfer_Overall_2Columns( _
    sourceWb As Workbook, _
    targetWb As Workbook, _
    rankingYear As String, _
    rankingName As String, _
    totalRespondents As Long _
) As Boolean

    On Error GoTo ErrorHandler
    Transfer_Overall_2Columns = False

    Dim sourceWs As Worksheet
    Dim sourceEvalWs As Worksheet
    Dim targetWs As Worksheet
    Dim i As Long

    Set sourceWs = sourceWb.Worksheets("総合対象企業")
    Set sourceEvalWs = sourceWb.Worksheets("評価項目")
    Set targetWs = targetWb.Worksheets("総合2つ")

    Module1_Main.LogMessage "  [2/3] 総合2つシートに転記中..."

    ' 【左側】総合ランキング
    ' タイトル (C4セル)
    Dim titleText As String
    titleText = "総合 (回答者数：" & Format(totalRespondents, "#,##0") & "名)"
    targetWs.Range("C4").Value = titleText

    ' データ転記 (D6:E10)
    Dim maxRows As Long
    maxRows = Application.Min(5, sourceWs.Cells(sourceWs.Rows.Count, 3).End(xlUp).Row - 4)

    For i = 1 To maxRows
        targetWs.Cells(5 + i, 4).Value = sourceWs.Cells(4 + i, 3).Value  ' 企業名
        targetWs.Cells(5 + i, 5).Value = Round(sourceWs.Cells(4 + i, 7).Value, 2)  ' 得点
    Next i

    ' 【右側】評価項目（例：加入手続き）
    ' タイトル (G4セル)
    Dim evalItemName As String
    evalItemName = sourceEvalWs.Cells(5, 1).Value  ' A5セルの評価項目名
    targetWs.Range("G4").Value = evalItemName

    ' データ転記 (H6:I10) - 評価項目の5-8行目
    For i = 1 To 4
        targetWs.Cells(5 + i, 8).Value = sourceEvalWs.Cells(4 + i, 4).Value  ' 企業名 (D列)
        targetWs.Cells(5 + i, 9).Value = Round(sourceEvalWs.Cells(4 + i, 5).Value, 2)  ' 得点 (E列)
    Next i

    Module1_Main.LogMessage "    [OK] 総合2つシート完了 (左:総合, 右:" & evalItemName & ")"
    Transfer_Overall_2Columns = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "    [ERROR] 総合2つシート転記エラー: " & Err.Description
    Transfer_Overall_2Columns = False
End Function

'========================================
' 総合ランキング転記: 総合+前回順位シート
'========================================
Private Function Transfer_Overall_WithPrevRank( _
    sourceWb As Workbook, _
    targetWb As Workbook, _
    rankingYear As String, _
    rankingName As String, _
    totalRespondents As Long _
) As Boolean

    On Error GoTo ErrorHandler
    Transfer_Overall_WithPrevRank = False

    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim i As Long

    Set sourceWs = sourceWb.Worksheets("総合対象企業")
    Set targetWs = targetWb.Worksheets("総合＋前回順位")

    Module1_Main.LogMessage "  [3/3] 総合＋前回順位シートに転記中..."

    ' 【左側】シンプル版 (B11:D14 - 4位まで)
    ' タイトル (C9セル)
    Dim titleText As String
    titleText = rankingName & " (回答者数：" & Format(totalRespondents, "#,##0") & "名)"
    targetWs.Range("C9").Value = titleText

    ' データ転記 (C11:D14)
    For i = 1 To 4
        targetWs.Cells(10 + i, 3).Value = sourceWs.Cells(4 + i, 3).Value  ' 企業名
        targetWs.Cells(10 + i, 4).Value = Round(sourceWs.Cells(4 + i, 7).Value, 2)  ' 得点
    Next i

    ' 【右側】前回順位付き (I11:J15 - 5位まで)
    ' タイトル (I9セル)
    targetWs.Range("I9").Value = titleText

    ' データ転記 (I11:J15)
    Dim maxRows As Long
    maxRows = Application.Min(5, sourceWs.Cells(sourceWs.Rows.Count, 3).End(xlUp).Row - 4)

    For i = 1 To maxRows
        targetWs.Cells(10 + i, 9).Value = sourceWs.Cells(4 + i, 3).Value  ' 企業名
        targetWs.Cells(10 + i, 10).Value = Round(sourceWs.Cells(4 + i, 7).Value, 2)  ' 得点
        ' 前回順位 (H列) は手動入力のため、転記しない
    Next i

    Module1_Main.LogMessage "    [OK] 総合＋前回順位シート完了"
    Transfer_Overall_WithPrevRank = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "    [ERROR] 総合＋前回順位シート転記エラー: " & Err.Description
    Transfer_Overall_WithPrevRank = False
End Function

'========================================
' 評価項目別ランキング転記: 評価・部門別シート
'========================================
Private Function Transfer_EvaluationItems( _
    sourceWb As Workbook, _
    targetWb As Workbook, _
    rankingYear As String, _
    rankingName As String _
) As Boolean

    On Error GoTo ErrorHandler
    Transfer_EvaluationItems = False

    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet

    Set sourceWs = sourceWb.Worksheets("評価項目")
    Set targetWs = targetWb.Worksheets("評価・部門別")

    Module1_Main.LogMessage "  評価項目別ランキング転記中..."

    ' タイトル設定 (B7セル)
    Dim titleText As String
    titleText = rankingYear & "年 オリコン顧客満足度(R)調査 " & rankingName & " 評価項目別ランキング"
    targetWs.Range("B7").Value = titleText

    ' 7つの評価項目の配置情報
    ' 配列: {評価項目名, ソース開始行, 転記先列(名前), 転記先列(順位), 転記先列(企業名), 転記先列(得点), 転記先行(データ開始)}
    Dim evalItems As Variant
    evalItems = Array( _
        Array("加入手続き", 5, 2, 2, 3, 4, 10), _
        Array("キャンペーン", 11, 6, 6, 7, 8, 10), _
        Array("初期設定のしやすさ", 17, 10, 10, 11, 12, 10), _
        Array("通信速度", 23, 14, 14, 15, 16, 10), _
        Array("料金プラン", 29, 2, 2, 3, 4, 16), _
        Array("端末のラインナップ", 35, 6, 6, 7, 8, 16), _
        Array("利用料金", 41, 10, 10, 11, 12, 16) _
    )

    Dim item As Variant
    Dim itemIndex As Long
    Dim sourceRow As Long
    Dim i As Long

    itemIndex = 0
    For Each item In evalItems
        itemIndex = itemIndex + 1

        Dim itemName As String
        Dim sourceStartRow As Long
        Dim nameCol As Long, rankCol As Long, companyCol As Long, scoreCol As Long
        Dim targetStartRow As Long

        itemName = item(0)
        sourceStartRow = item(1)
        nameCol = item(2)
        rankCol = item(3)
        companyCol = item(4)
        scoreCol = item(5)
        targetStartRow = item(6)

        ' 評価項目名を設定
        targetWs.Cells(targetStartRow - 2, nameCol).Value = itemName

        ' データ転記 (3位まで)
        For i = 0 To 2
            sourceRow = sourceStartRow + i

            ' 企業名転記 (D列 → 転記先企業名列)
            targetWs.Cells(targetStartRow + i, companyCol).Value = sourceWs.Cells(sourceRow, 4).Value

            ' 得点転記 (E列 → 転記先得点列) - 小数点以下2桁
            targetWs.Cells(targetStartRow + i, scoreCol).Value = Round(sourceWs.Cells(sourceRow, 5).Value, 2)
        Next i

        Module1_Main.LogMessage "    [" & itemIndex & "/7] " & itemName & " 転記完了"
    Next item

    Module1_Main.LogMessage "  [OK] 評価項目7つ全て転記完了"
    Transfer_EvaluationItems = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "    [ERROR] 評価項目転記エラー: " & Err.Description
    Transfer_EvaluationItems = False
End Function
