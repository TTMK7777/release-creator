# Claude Code実装ガイド - タスクリスト

**対象:** Claude Code実装者  
**難易度:** 中級  
**推定実装時間:** 3-5時間  
**言語:** VBA (Visual Basic for Applications)

---

## 🎯 実装の目標

完全オフラインで動作するプレスリリース自動生成ツールをVBAで実装する。

---

## 📁 提供されているファイル

### 入力ファイル(サンプル)
1. `_資料_携帯キャリア_ランキング結果2024.xlsx` - 元データ
2. `2025年リフォーム_リリース内表.xlsx` - リリース用表テンプレート
3. `2023年12月発表__携帯キャリア_..._ニュースリリース__.docx` - Wordテンプレート

### ドキュメント
1. `プレスリリース自動生成_Claude Code引き継ぎ資料.md` - 全体概要
2. `技術仕様書_データ構造とマッピング.md` - 詳細仕様

---

## 🚀 クイックスタート

### ステップ0: 環境準備

```vba
' 1. 新しいExcelファイルを作成
' 2. Alt+F11でVBEを開く
' 3. ツール > 参照設定 > 以下をチェック:
'    ☑ Microsoft Excel XX.0 Object Library
'    ☑ Microsoft Word XX.0 Object Library
```

### ステップ1: モジュール作成

VBEで以下のモジュールを作成:
- `Module1_Main` - メイン制御
- `Module2_Data` - データ処理
- `Module3_Image` - 画像生成
- `Module4_Word` - Word操作
- `Module5_Utils` - ユーティリティ

---

## 📝 実装タスクリスト

### Phase 1: 基礎実装(1-2時間)

#### Task 1.1: メインマクロの骨組み
```vba
' Module1_Main

Option Explicit

Sub プレスリリース自動生成()
    ' TODO: 実装
    ' 1. ファイルパス設定
    ' 2. データ転記呼び出し
    ' 3. 画像生成呼び出し
    ' 4. Word生成呼び出し
    ' 5. 完了メッセージ
End Sub
```

**チェックポイント:**
- [ ] マクロが実行できる
- [ ] エラーハンドリングが動作する
- [ ] 実行時間が計測できる

---

#### Task 1.2: ファイルパス設定

```vba
' Module1_Main

Private Function ファイルパス取得() As Dictionary
    Dim paths As Dictionary
    Set paths = New Dictionary
    
    ' TODO: 実装
    paths.Add "元データ", ThisWorkbook.Path & "\_資料_携帯キャリア_ランキング結果2024.xlsx"
    paths.Add "リリース表", ThisWorkbook.Path & "\リリース内表.xlsx"
    paths.Add "テンプレート", ThisWorkbook.Path & "\テンプレート.docx"
    paths.Add "出力フォルダ", ThisWorkbook.Path & "\output\"
    
    Set ファイルパス取得 = paths
End Function
```

**チェックポイント:**
- [ ] ファイルパスが正しく取得できる
- [ ] 出力フォルダが自動作成される

---

#### Task 1.3: ファイル存在チェック

```vba
' Module5_Utils

Function ファイル存在チェック(filePath As String) As Boolean
    ' TODO: 実装
    ファイル存在チェック = (Dir(filePath) <> "")
End Function

Function シート存在チェック(wb As Workbook, sheetName As String) As Boolean
    ' TODO: 実装
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    シート存在チェック = Not ws Is Nothing
    On Error GoTo 0
End Function
```

**チェックポイント:**
- [ ] 存在するファイルでTrueが返る
- [ ] 存在しないファイルでFalseが返る

---

### Phase 2: データ処理(1-2時間)

#### Task 2.1: 総合ランキング転記

```vba
' Module2_Data

Sub データ転記(元DataPath As String, リリース表Path As String)
    
    Dim 元wb As Workbook
    Dim リリースwb As Workbook
    
    ' TODO: ファイルオープン
    Set 元wb = Workbooks.Open(元DataPath, ReadOnly:=True)
    Set リリースwb = Workbooks.Open(リリース表Path)
    
    ' TODO: 総合ランキング転記
    Call 総合ランキング転記(元wb.Worksheets("総合対象企業"), _
                           リリースwb.Worksheets("総合3つ"))
    
    ' TODO: 保存・クローズ
    リリースwb.Save
    元wb.Close SaveChanges:=False
    リリースwb.Close SaveChanges:=True
    
End Sub

Private Sub 総合ランキング転記(元ws As Worksheet, リリースws As Worksheet)
    
    ' TODO: 実装
    ' 元データ5-7行目 → リリース表9-11行目
    
    Dim i As Long
    Dim 転記先Row As Long
    転記先Row = 9
    
    For i = 5 To 7
        With リリースws
            .Cells(転記先Row, 2).Value = (i - 4) & "位"
            .Cells(転記先Row, 3).Value = 元ws.Cells(i, 3).Value
            .Cells(転記先Row, 4).Value = Round(元ws.Cells(i, 7).Value, 1)
        End With
        転記先Row = 転記先Row + 1
    Next i
    
End Sub
```

**テストケース:**
```vba
Sub Test_総合ランキング転記()
    ' TODO: テストコード
    ' 期待値: リリース表のB9に"1位"、C9に"楽天モバイル"、D9に69.5
End Sub
```

**チェックポイント:**
- [ ] 1位のデータが正しく転記される
- [ ] 2位、3位も正しく転記される
- [ ] 得点が小数点1桁で丸められる

---

#### Task 2.2: 評価項目別転記

```vba
' Module2_Data

Private Sub 評価項目別転記(元ws As Worksheet, リリースws As Worksheet)
    
    ' TODO: 実装
    ' 評価項目リスト
    Dim 項目リスト As Variant
    項目リスト = Array("加入手続き", "キャンペーン", "初期設定のしやすさ", _
                      "通信速度", "料金プラン", "端末のラインナップ", _
                      "利用料金", "サポートサービス", "付帯サービス")
    
    Dim 項目 As Variant
    Dim リリース列 As Long
    リリース列 = 2
    
    For Each 項目 In 項目リスト
        ' 該当項目の1位を検索
        Dim row As Long
        For row = 5 To 56
            If 元ws.Cells(row, 1).Value = 項目 And _
               元ws.Cells(row, 2).Value = 1 Then
                
                ' リリース表に書き込み
                リリースws.Cells(10, リリース列).Value = "1位"
                リリースws.Cells(11, リリース列).Value = 元ws.Cells(row, 4).Value
                リリースws.Cells(12, リリース列).Value = Round(元ws.Cells(row, 5).Value, 1)
                
                Exit For
            End If
        Next row
        
        リリース列 = リリース列 + 4
    Next 項目
    
End Sub
```

**チェックポイント:**
- [ ] 9つの評価項目すべてが転記される
- [ ] 各項目の1位企業が正しい
- [ ] 得点が正しく丸められる

---

### Phase 3: 画像生成(30分-1時間)

#### Task 3.1: 総合ランキング画像化

```vba
' Module3_Image

Sub 表を画像化(リリース表Path As String, 出力フォルダPath As String)
    
    Dim wb As Workbook
    Set wb = Workbooks.Open(リリース表Path, ReadOnly:=True)
    
    ' TODO: 総合ランキング画像化
    Call 総合ランキング画像化(wb, 出力フォルダPath)
    
    ' TODO: 評価項目別画像化
    Call 評価項目別画像化(wb, 出力フォルダPath)
    
    wb.Close SaveChanges:=False
    
End Sub

Private Sub 総合ランキング画像化(wb As Workbook, 出力フォルダ As String)
    
    Dim ws As Worksheet
    Set ws = wb.Worksheets("総合3つ")
    
    ' TODO: 実装
    Dim rng As Range
    Set rng = ws.Range("B7:D20")
    
    rng.CopyPicture Appearance:=xlScreen, Format:=xlPicture
    
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add( _
        Left:=rng.Left, Top:=rng.Top, _
        Width:=rng.Width, Height:=rng.Height)
    
    With cht.Chart
        .Paste
        .Export Filename:=出力フォルダ & "総合ランキング.png", FilterName:="PNG"
    End With
    
    cht.Delete
    
End Sub
```

**チェックポイント:**
- [ ] PNG画像が生成される
- [ ] 画像が鮮明に表示される
- [ ] ファイルサイズが適切(100KB-1MB程度)

---

#### Task 3.2: 評価項目別画像化

```vba
' Module3_Image

Private Sub 評価項目別画像化(wb As Workbook, 出力フォルダ As String)
    
    ' TODO: 実装(総合ランキング画像化と同様の処理)
    Dim ws As Worksheet
    Set ws = wb.Worksheets("フルリフォーム 評価項目別")
    
    Dim rng As Range
    Set rng = ws.Range("B8:X30")
    
    ' 以下、同様の処理...
    
End Sub
```

**チェックポイント:**
- [ ] 評価項目表が画像化される
- [ ] 表全体が含まれている

---

### Phase 4: Word操作(1-2時間)

#### Task 4.1: Word起動と基本操作

```vba
' Module4_Word

Sub Word生成(テンプレートPath As String, 出力フォルダPath As String, _
             リリース表Path As String)
    
    Dim wdApp As Object
    Dim wdDoc As Object
    
    ' TODO: Word起動
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    wdApp.Visible = False
    
    ' TODO: テンプレート読込
    Set wdDoc = wdApp.Documents.Open(テンプレートPath)
    
    ' TODO: データ取得
    Dim リリースwb As Workbook
    Set リリースwb = Workbooks.Open(リリース表Path, ReadOnly:=True)
    
    ' TODO: 日付更新
    Call Word日付更新(wdDoc)
    
    ' TODO: タイトル更新
    Call Wordタイトル更新(wdDoc, リリースwb)
    
    ' TODO: 本文更新
    Call Word本文更新(wdDoc, リリースwb)
    
    ' TODO: 画像差し替え
    Call Word画像差し替え(wdDoc, 出力フォルダPath & "総合ランキング.png", 1)
    Call Word画像差し替え(wdDoc, 出力フォルダPath & "評価項目別ランキング.png", 2)
    
    ' TODO: 保存
    Dim 出力ファイル名 As String
    出力ファイル名 = 出力フォルダPath & "プレスリリース_" & _
                     Format(Now, "yyyymmdd") & ".docx"
    wdDoc.SaveAs2 Filename:=出力ファイル名
    
    ' TODO: クリーンアップ
    wdDoc.Close
    wdApp.Quit
    リリースwb.Close SaveChanges:=False
    
End Sub
```

**チェックポイント:**
- [ ] Wordが起動する
- [ ] テンプレートが読み込まれる
- [ ] ファイルが保存される

---

#### Task 4.2: 日付更新

```vba
' Module4_Word

Private Sub Word日付更新(wdDoc As Object)
    
    ' TODO: 2024年 → 2025年
    With wdDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "2024年"
        .Replacement.Text = "2025年"
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    ' TODO: 2023年 → 2024年
    With wdDoc.Content.Find
        .Text = "2023年"
        .Replacement.Text = "2024年"
        .Execute Replace:=2
    End With
    
    ' TODO: 発表日更新
    Dim 今日 As String
    今日 = Format(Now, "yyyy年m月d日") & "(" & Format(Now, "aaa") & ")"
    
    With wdDoc.Content.Find
        .Text = "202[0-9]年[0-9]{1,2}月[0-9]{1,2}日"
        .MatchWildcards = True
        .Replacement.Text = 今日
        .Execute Replace:=2
    End With
    
End Sub
```

**チェックポイント:**
- [ ] 年号が更新される
- [ ] 今日の日付になる
- [ ] 曜日が正しい

---

#### Task 4.3: タイトル更新

```vba
' Module4_Word

Private Sub Wordタイトル更新(wdDoc As Object, リリースwb As Workbook)
    
    Dim ws As Worksheet
    Set ws = リリースwb.Worksheets("総合3つ")
    
    ' TODO: 1位企業名取得
    Dim 一位企業 As String
    一位企業 = ws.Cells(9, 3).Value
    
    ' TODO: タイトル生成(仮: 3年連続)
    Dim 新タイトル As String
    新タイトル = "【" & 一位企業 & "】が3年連続総合1位に"
    
    ' TODO: Wordのタイトル部分を探して置換
    Dim para As Object
    For Each para In wdDoc.Paragraphs
        If InStr(para.Range.Text, "【") > 0 And _
           InStr(para.Range.Text, "総合1位") > 0 Then
            para.Range.Text = 新タイトル
            Exit For
        End If
    Next para
    
End Sub
```

**チェックポイント:**
- [ ] タイトルが正しく生成される
- [ ] Word内で正しい位置に挿入される

---

#### Task 4.4: 本文更新

```vba
' Module4_Word

Private Sub Word本文更新(wdDoc As Object, リリースwb As Workbook)
    
    Dim ws As Worksheet
    Set ws = リリースwb.Worksheets("総合3つ")
    
    ' TODO: データ取得
    Dim 一位企業 As String
    Dim 二位企業 As String
    Dim 三位企業 As String
    
    一位企業 = ws.Cells(9, 3).Value
    二位企業 = ws.Cells(10, 3).Value
    三位企業 = ws.Cells(11, 3).Value
    
    ' TODO: 本文生成
    Dim 本文 As String
    本文 = 本文生成(一位企業, 二位企業, 三位企業)
    
    ' TODO: Word内の本文部分を探して置換
    ' (実装方法は複数あるため、テンプレート構造に応じて調整)
    
End Sub

Private Function 本文生成(一位 As String, 二位 As String, 三位 As String) As String
    
    Dim text As String
    
    text = "オリコン株式会社(本社:東京都港区 代表取締役社長:小池 恒)は、" & vbCrLf
    text = text & "年1回の満足度調査「オリコン顧客満足度®調査」を実施し、" & vbCrLf
    text = text & Format(Now, "yyyy年m月d日") & "(火) 14時に" & vbCrLf
    text = text & "調査結果を公式サイト内にて発表しました。" & vbCrLf
    text = text & vbCrLf
    text = text & "ランキングの結果は、下記の通りとなりました。" & vbCrLf
    text = text & vbCrLf
    text = text & "【TOPICS】" & vbCrLf
    text = text & "■【" & 一位 & "】が3年連続総合1位に" & vbCrLf
    text = text & "■【" & 二位 & "】が総合2位" & vbCrLf
    text = text & "■【" & 三位 & "】が総合3位" & vbCrLf
    
    本文生成 = text
    
End Function
```

**チェックポイント:**
- [ ] 本文が生成される
- [ ] TOPICSが含まれる
- [ ] 企業名が正しい

---

#### Task 4.5: 画像差し替え

```vba
' Module4_Word

Private Sub Word画像差し替え(wdDoc As Object, 画像Path As String, 画像番号 As Integer)
    
    On Error Resume Next
    
    ' TODO: 既存画像を削除
    If wdDoc.InlineShapes.Count >= 画像番号 Then
        wdDoc.InlineShapes(画像番号).Delete
    End If
    
    On Error GoTo 0
    
    ' TODO: 新しい画像を挿入
    ' (挿入位置は実際のテンプレート構造に応じて調整)
    
    Dim 挿入位置 As Object
    Set 挿入位置 = wdDoc.Paragraphs(10 + 画像番号 * 5).Range
    
    挿入位置.InlineShapes.AddPicture _
        Filename:=画像Path, _
        LinkToFile:=False, _
        SaveWithDocument:=True
    
End Sub
```

**チェックポイント:**
- [ ] 画像が挿入される
- [ ] 画像が正しい位置にある
- [ ] 画像サイズが適切

---

### Phase 5: エラーハンドリングとテスト(30分-1時間)

#### Task 5.1: エラーハンドリング実装

```vba
' Module1_Main

Sub プレスリリース自動生成()
    On Error GoTo ErrorHandler
    
    ' メイン処理...
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "エラーが発生しました: " & vbNewLine & _
           "エラー番号: " & Err.Number & vbNewLine & _
           "説明: " & Err.Description, vbCritical
    
    Call エラーログ出力(Err.Description)
    
End Sub
```

**チェックポイント:**
- [ ] エラー時にメッセージが表示される
- [ ] エラーログが記録される
- [ ] 画面更新が元に戻る

---

#### Task 5.2: 統合テスト

```vba
' テストモジュール

Sub Test_フルフロー()
    
    ' TODO: サンプルデータで全工程をテスト
    
    Debug.Print "=== フルフローテスト開始 ==="
    
    Call プレスリリース自動生成
    
    ' TODO: 結果確認
    ' - output/総合ランキング.pngが存在するか
    ' - output/評価項目別ランキング.pngが存在するか
    ' - output/プレスリリース_YYYYMMDD.docxが存在するか
    
    Debug.Print "=== フルフローテスト完了 ==="
    
End Sub
```

**チェックポイント:**
- [ ] 全工程が正常に完了する
- [ ] 5分以内に完了する
- [ ] 生成されたファイルが正しい

---

## 🐛 デバッグTips

### デバッグモードでステップ実行

```vba
Sub デバッグ実行_データ転記のみ()
    Dim 元Path As String
    Dim リリースPath As String
    
    元Path = "C:\...\元データ.xlsx"
    リリースPath = "C:\...\リリース表.xlsx"
    
    Call データ転記(元Path, リリースPath)
    
    ' 結果確認
    Dim wb As Workbook
    Set wb = Workbooks.Open(リリースPath)
    Debug.Print "1位企業: " & wb.Worksheets("総合3つ").Cells(9, 3).Value
    wb.Close False
End Sub
```

### イミディエイトウィンドウでの確認

```vba
' イミディエイトウィンドウ(Ctrl+G)で実行:
? Workbooks("元データ.xlsx").Worksheets("総合対象企業").Cells(5, 3).Value
→ "楽天モバイル"

? Round(69.520559, 1)
→ 69.5
```

---

## ⚠️ 重要な注意事項

### セキュリティ
- **機密情報を外部に送信しないこと**
- **APIを使用しないこと**
- **完全オフラインで動作すること**

### パフォーマンス
- `Application.ScreenUpdating = False` を使用
- オブジェクトは適切に解放
- 不要なループを避ける

### 保守性
- 変数名は日本語OK(わかりやすさ優先)
- コメントは日本語で詳しく
- マジックナンバーは定数化

---

## 📊 完成度チェックリスト

### 機能実装
- [ ] データ転記(総合ランキング)
- [ ] データ転記(評価項目別)
- [ ] 画像生成(総合ランキング)
- [ ] 画像生成(評価項目別)
- [ ] Word生成(日付更新)
- [ ] Word生成(タイトル更新)
- [ ] Word生成(本文更新)
- [ ] Word生成(画像差し替え)

### エラーハンドリング
- [ ] ファイル存在チェック
- [ ] シート存在チェック
- [ ] データ整合性チェック
- [ ] エラーログ出力

### テスト
- [ ] 単体テスト(各関数)
- [ ] 統合テスト(フルフロー)
- [ ] 異常系テスト

### ドキュメント
- [ ] コード内コメント
- [ ] 使用方法のREADME
- [ ] 既知の問題・制限事項

---

## 🎓 学習リソース

### VBA基礎
- [Excel VBA リファレンス(Microsoft)](https://docs.microsoft.com/ja-jp/office/vba/api/overview/excel)
- [Word VBA リファレンス(Microsoft)](https://docs.microsoft.com/ja-jp/office/vba/api/overview/word)

### トラブルシューティング
- エラー番号で検索
- StackOverflowで類似事例を探す
- Microsoft Docsの公式ドキュメント参照

---

## 💬 質問・サポート

わからないことがあれば、以下の情報を添えて質問してください:
1. どのTask番号で詰まったか
2. エラーメッセージ(あれば)
3. 試したこと
4. 期待する動作

**実装頑張ってください!🚀**
