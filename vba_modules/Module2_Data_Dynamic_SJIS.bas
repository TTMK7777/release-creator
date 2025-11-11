Attribute VB_Name = "Module2_Data"
'========================================
' Module2_Data_Dynamic: 動的転記モジュール
' 作成日: 2025年11月11日
' バージョン: 6.0 (完全動的版 - ハードコーディングなし)
'========================================
Option Explicit

' 評価項目情報を格納する構造体
Private Type EvaluationItem
    Name As String
    StartRow As Long
    EndRow As Long
    CompanyCount As Long
End Type

'========================================
' メインデータ転記関数（完全動的版）
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
    Module1_Main.LogMessage "動的データ転記を開始"
    Module1_Main.LogMessage "======================================"
    Module1_Main.LogMessage "  元データ: " & sourceFilePath
    Module1_Main.LogMessage "  転記先: " & targetFilePath
    Module1_Main.LogMessage ""

    ' ファイルを開く
    Set sourceWb = Workbooks.Open(sourceFilePath, ReadOnly:=True)
    Set targetWb = Workbooks.Open(targetFilePath)

    ' デバッグ: ターゲットファイルのシート名を全て表示
    Module1_Main.LogMessage "  [DEBUG] ターゲットファイルのシート一覧:"
    Dim ws As Worksheet
    For Each ws In targetWb.Worksheets
        Module1_Main.LogMessage "    [" & ws.Index & "] " & ws.Name
    Next ws
    Module1_Main.LogMessage ""

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
    ' Phase 2: 評価項目別ランキング転記（完全動的）
    ' ================================================
    Module1_Main.LogMessage "[Phase 2] 評価項目別ランキング転記..."

    If Not Transfer_EvaluationItems_Dynamic(sourceWb, targetWb, rankingYear, rankingName) Then
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
' 評価項目の自動検出
'========================================
Private Function DetectEvaluationItems(sourceWs As Worksheet) As EvaluationItem()
    On Error GoTo ErrorHandler

    Dim items() As EvaluationItem
    Dim itemCount As Long
    Dim currentRow As Long
    Dim lastRow As Long
    Dim prevItemName As String
    Dim currentItemName As String
    Dim itemStartRow As Long

    itemCount = 0
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
    prevItemName = ""
    itemStartRow = 0

    Module1_Main.LogMessage "  [自動検出] 評価項目を検出中..."

    ' A列をスキャンして評価項目の切り替わりを検出
    For currentRow = 5 To lastRow
        currentItemName = Trim(sourceWs.Cells(currentRow, 1).Value)

        ' 評価項目名が変わった = 新しい評価項目の開始
        If currentItemName <> "" And currentItemName <> prevItemName Then
            ' 前の評価項目を配列に追加
            If prevItemName <> "" And itemStartRow > 0 Then
                itemCount = itemCount + 1
                ReDim Preserve items(1 To itemCount)

                items(itemCount).Name = prevItemName
                items(itemCount).StartRow = itemStartRow
                items(itemCount).EndRow = currentRow - 1
                items(itemCount).CompanyCount = (currentRow - 1) - itemStartRow + 1

                Module1_Main.LogMessage "    [" & itemCount & "] " & prevItemName & _
                                        " (行" & itemStartRow & "-" & (currentRow - 1) & ", " & _
                                        items(itemCount).CompanyCount & "社)"
            End If

            ' 新しい評価項目の開始
            itemStartRow = currentRow
            prevItemName = currentItemName
        End If
    Next currentRow

    ' 最後の評価項目を追加
    If prevItemName <> "" And itemStartRow > 0 Then
        itemCount = itemCount + 1
        ReDim Preserve items(1 To itemCount)

        items(itemCount).Name = prevItemName
        items(itemCount).StartRow = itemStartRow
        items(itemCount).EndRow = lastRow
        items(itemCount).CompanyCount = lastRow - itemStartRow + 1

        Module1_Main.LogMessage "    [" & itemCount & "] " & prevItemName & _
                                " (行" & itemStartRow & "-" & lastRow & ", " & _
                                items(itemCount).CompanyCount & "社)"
    End If

    Module1_Main.LogMessage "  [OK] 評価項目検出完了: " & itemCount & "項目"

    DetectEvaluationItems = items
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "  [ERROR] 評価項目検出エラー: " & Err.Description
    ReDim items(0 To 0)
    DetectEvaluationItems = items
End Function

'========================================
' 評価項目別ランキング転記（完全動的版）
'========================================
Private Function Transfer_EvaluationItems_Dynamic( _
    sourceWb As Workbook, _
    targetWb As Workbook, _
    rankingYear As String, _
    rankingName As String _
) As Boolean

    On Error GoTo ErrorHandler
    Transfer_EvaluationItems_Dynamic = False

    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet

    Set sourceWs = sourceWb.Worksheets("評価項目")
    Set targetWs = targetWb.Worksheets("評価・部門別")

    Module1_Main.LogMessage "  評価項目別ランキング転記中（動的検出）..."

    ' タイトル設定 (B7セル)
    Dim titleText As String
    titleText = rankingYear & "年 オリコン顧客満足度(R)調査 " & rankingName & " 評価項目別ランキング"
    targetWs.Range("B7").Value = titleText

    ' 評価項目を動的に検出
    Dim items() As EvaluationItem
    items = DetectEvaluationItems(sourceWs)

    If UBound(items) = 0 Then
        Module1_Main.LogMessage "  [WARN] 評価項目が検出されませんでした"
        Transfer_EvaluationItems_Dynamic = True
        Exit Function
    End If

    ' 転記先の配置パターン（テンプレート側の固定レイアウト）
    ' 上段4カラム + 下段3カラム = 最大7項目
    Dim targetColumns As Variant
    targetColumns = Array( _
        Array(2, 10), _
        Array(6, 10), _
        Array(10, 10), _
        Array(14, 10), _
        Array(2, 16), _
        Array(6, 16), _
        Array(10, 16) _
    )

    ' 各評価項目を転記
    Dim i As Long
    Dim itemIndex As Long
    Dim maxItems As Long
    Dim nameCol As Long, dataStartRow As Long
    Dim sourceRow As Long
    Dim targetRow As Long

    maxItems = Application.Min(UBound(items), UBound(targetColumns) + 1)

    For itemIndex = 1 To maxItems
        nameCol = targetColumns(itemIndex - 1)(0)        ' 転記先の列（企業名列の-1）
        dataStartRow = targetColumns(itemIndex - 1)(1)   ' 転記先のデータ開始行

        ' 評価項目名を設定（データ開始行の-2行目）
        targetWs.Cells(dataStartRow - 2, nameCol).Value = items(itemIndex).Name

        ' ランキングデータを転記（最大3位まで）
        Dim maxRank As Long
        maxRank = Application.Min(3, items(itemIndex).CompanyCount)

        For i = 0 To maxRank - 1
            sourceRow = items(itemIndex).StartRow + i
            targetRow = dataStartRow + i

            ' 企業名転記 (D列 → 転記先企業名列)
            targetWs.Cells(targetRow, nameCol + 1).Value = sourceWs.Cells(sourceRow, 4).Value

            ' 得点転記 (E列 → 転記先得点列) - 小数点以下2桁
            targetWs.Cells(targetRow, nameCol + 2).Value = Round(sourceWs.Cells(sourceRow, 5).Value, 2)
        Next i

        Module1_Main.LogMessage "    [" & itemIndex & "/" & maxItems & "] " & _
                                items(itemIndex).Name & " 転記完了 (" & maxRank & "位まで)"
    Next itemIndex

    ' 検出項目数が7より多い場合の警告
    If UBound(items) > 7 Then
        Module1_Main.LogMessage "  [WARN] 評価項目が" & UBound(items) & "個検出されましたが、" & _
                                "テンプレートは7項目までしか対応していません"
    End If

    Module1_Main.LogMessage "  [OK] 評価項目転記完了: " & maxItems & "項目"
    Transfer_EvaluationItems_Dynamic = True
    Exit Function

ErrorHandler:
    Module1_Main.LogMessage "    [ERROR] 評価項目転記エラー: " & Err.Description
    Transfer_EvaluationItems_Dynamic = False
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

    ' データ件数を動的に検出
    Dim maxRows As Long
    maxRows = Application.Min(5, sourceWs.Cells(sourceWs.Rows.Count, 3).End(xlUp).Row - 4)

    ' データ転記 (D6:E10 - 最大5位まで)
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

    ' データ件数を動的に検出
    Dim maxRows As Long
    maxRows = Application.Min(5, sourceWs.Cells(sourceWs.Rows.Count, 3).End(xlUp).Row - 4)

    ' データ転記 (D6:E10)
    For i = 1 To maxRows
        targetWs.Cells(5 + i, 4).Value = sourceWs.Cells(4 + i, 3).Value  ' 企業名
        targetWs.Cells(5 + i, 5).Value = Round(sourceWs.Cells(4 + i, 7).Value, 2)  ' 得点
    Next i

    ' 【右側】評価項目（最初の評価項目を自動検出）
    ' 評価項目名を取得 (A列の最初の非空白セル)
    Dim evalItemName As String
    Dim evalStartRow As Long
    evalStartRow = 5  ' 通常は5行目から開始
    evalItemName = sourceEvalWs.Cells(evalStartRow, 1).Value

    ' タイトル (G4セル)
    targetWs.Range("G4").Value = evalItemName

    ' データ転記 (H6:I10) - 最大5社まで
    Dim evalMaxRows As Long
    evalMaxRows = 0

    ' 同じ評価項目名が続く行数をカウント
    Dim checkRow As Long
    For checkRow = evalStartRow To sourceEvalWs.Cells(sourceEvalWs.Rows.Count, 1).End(xlUp).Row
        If sourceEvalWs.Cells(checkRow, 1).Value = evalItemName Then
            evalMaxRows = evalMaxRows + 1
        Else
            Exit For
        End If
    Next checkRow

    evalMaxRows = Application.Min(5, evalMaxRows)

    For i = 1 To evalMaxRows
        targetWs.Cells(5 + i, 8).Value = sourceEvalWs.Cells(evalStartRow - 1 + i, 4).Value  ' 企業名 (D列)
        targetWs.Cells(5 + i, 9).Value = Round(sourceEvalWs.Cells(evalStartRow - 1 + i, 5).Value, 2)  ' 得点 (E列)
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

    ' シート名の存在確認
    On Error Resume Next
    Set targetWs = targetWb.Worksheets("総合＋前回順位")
    If targetWs Is Nothing Then
        ' 半角+で試す
        Set targetWs = targetWb.Worksheets("総合+前回順位")
    End If
    If targetWs Is Nothing Then
        Module1_Main.LogMessage "  [SKIP] 総合＋前回順位シートが見つかりません（スキップ）"
        Transfer_Overall_WithPrevRank = True
        Exit Function
    End If
    On Error GoTo ErrorHandler

    Module1_Main.LogMessage "  [3/3] " & targetWs.Name & "シートに転記中..."

    ' タイトル生成
    Dim titleText As String
    titleText = rankingName & " (回答者数：" & Format(totalRespondents, "#,##0") & "名)"

    ' 【左側】シンプル版 (C11:D14 - 最大4位まで)
    ' タイトル (C9セル)
    targetWs.Range("C9").Value = titleText

    ' データ件数を動的に検出
    Dim maxRowsLeft As Long
    maxRowsLeft = Application.Min(4, sourceWs.Cells(sourceWs.Rows.Count, 3).End(xlUp).Row - 4)

    ' データ転記 (C11:D14)
    For i = 1 To maxRowsLeft
        targetWs.Cells(10 + i, 3).Value = sourceWs.Cells(4 + i, 3).Value  ' 企業名
        targetWs.Cells(10 + i, 4).Value = Round(sourceWs.Cells(4 + i, 7).Value, 2)  ' 得点
    Next i

    ' 【右側】前回順位付き (I11:J15 - 最大5位まで)
    ' タイトル (I9セル)
    targetWs.Range("I9").Value = titleText

    ' データ件数を動的に検出
    Dim maxRowsRight As Long
    maxRowsRight = Application.Min(5, sourceWs.Cells(sourceWs.Rows.Count, 3).End(xlUp).Row - 4)

    ' データ転記 (I11:J15)
    For i = 1 To maxRowsRight
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
