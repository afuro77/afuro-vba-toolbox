# Afuro VBA Toolbox

afuro77用の「共通部品をごった煮で揃えたVBAの道具箱」です。配布物は標準モジュール `.bas` とクラス `.cls`。ツール用Excelブックや自動実行イベントはありません。

最初に **[部品一覧.html](部品一覧.html)** を開き、「罫線」「コード」「CSV」などで検索してください。各部品に日本語の用途、使用例、コード、依存関係を載せています。[Markdown版](部品一覧.md) もあります。

## 使い方

1. ZIPを展開します。Excelブックへ組み込む場合はAlt+F11でVBEを開きます。
2. 「ファイル → ファイルのインポート」で `vbe/modules/mdlAT_*.bas` と、必要な `vbe/classes/clsAT_*.cls` を取り込みます。日本語Windows向けCP932・CRLFで変換済みです。
3. VBEの「デバッグ → VBAProjectのコンパイル」を実行します。
4. 部品一覧の使用例を自分の処理に組み込みます。`ws`・`source`・`target`・`path`などは、自分で宣言・設定する変数です。

全部入りで使うなら、`modules`の11ファイルと`classes`の3ファイルをすべて取り込めます。標準モジュールを1ファイルにしたい場合だけ、11ファイルの代わりに `vbe/all-in-one/mdlAT_All.bas` を使います。**分割版と全部入り版を同時に取り込まないでください。** クラスはどちらの場合も別ファイルです。

VBAの標準モジュールのエクスポート拡張子は `.bas` です。`mdl` はモジュール名の接頭辞として使っています。クラスの拡張子は `.cls` です。

`src` は編集・差分管理用のUTF-8ソース、`vbe` はVBEへのインポート用です。通常は変換作業やPythonのインストールは不要です。CSVのUTF-8ファイル操作とファイル列挙はWindowsのADODB.Stream／FileSystemObjectを遅延バインディングで利用します。参照設定の追加やWindows API宣言はありません。

## 何が入っているか

| モジュール | 内容 | 他モジュールへの依存 |
| --- | --- | --- |
| mdlAT_Excel | 設定退避・復元、シート取得・名前処理、進捗 | なし |
| mdlAT_Range | 最終行・列、検索、配列転記、範囲結合、空行削除、行コピー、ソート | なし |
| mdlAT_Text | 空値処理、文字列分割・置換、コード名称、列名変換、CSVエスケープ | なし |
| mdlAT_Format | 配置、外枠・格子、塗りつぶし、フォント、表示形式、丸め | なし |
| mdlAT_Files | パス、ファイル列挙・コピー、テキスト入出力、選択ダイアログ | なし |
| mdlAT_Array | 配列変換、重複除去、空値除去、転置、列・行抽出 | なし |
| mdlAT_Dictionary | 辞書、件数集計、キー行索引、キー別合計 | なし |
| mdlAT_Date | 厳密な日付変換、月初・月末・週、日時文字列 | なし |
| mdlAT_Validation | 必須、重複、数値範囲、入力規則、ロック | なし |
| mdlAT_Shapes | 吹き出し転記、メモ、接頭辞による図形削除 | なし |
| mdlAT_Csv | CSV解析・生成、範囲のCSV保存 | Text・Range・Files |
| clsAT_AppGuard | 設定退避と後片付けをオブジェクトで管理 | なし |
| clsAT_Rule | 比較条件を設定してセル・値をチェック | なし |
| clsAT_Logger | 指定シートへ日時・処理名・エラーを記録 | なし |

同じモジュール内の補助関数はまとめてコピーしてください。単一関数を抜き出す場合は、その関数が呼ぶ `AT_` 関数も必要です。Excel設定の管理は標準モジュール方式とクラス方式のどちらかを選べます。

## よく使う形

```vb
Dim ws As Worksheet, last As Long, data As Variant
Set ws = ThisWorkbook.Worksheets("データ")
last = AT_LastRow(ws, "A")
If last >= 2 Then data = AT_ReadMatrix(ws.Range("A2:C" & last))
```

```vb
Dim guard As New clsAT_AppGuard
Dim errorNumber As Long, errorMessage As String
On Error GoTo Failed
guard.Start
' 自分の処理
guard.Restore
Exit Sub
Failed:
errorNumber = Err.Number: errorMessage = Err.Description
On Error Resume Next
guard.Restore
On Error GoTo 0
Err.Raise errorNumber, "MyMacro", errorMessage
```

エラー処理の例は自分のSubの中へ貼り付けます。VBEのリセットや `End` 文ではClass_Terminateが働かないため、Restoreを明示的に呼ぶ形を基本にします。設定の入れ子は開始と逆順に復元してください。

複数部品を組み合わせた集計、CSV取込、必須欄の色付け例は `examples/mdlAT_Recipes.bas` にあります。

## 共通の扱い

- Worksheet・Workbook・Rangeは引数で指定します。ActiveSheet・Select・固定の個人パスを使いません。
- 行・列番号と件数はLong。1セルの読み込みでも2次元配列を返します。
- 空値と0、False、エラーを区別します。型不正や不明なオプションは知らせます。
- 検索で見つからない場合、セル／シートはNothing、列番号は0です。配列抽出・空CSV・空辞書の配列化はEmptyを返すものがあります。
- ファイルコピー・CSV保存・UTF-8出力は既定で上書きしません。必要なときだけoverwriteをTrueにします。
- 空行削除は範囲内が空の行を、**シート全体の行として**削除します。範囲外の列にデータがある場合は使わないでください。
- CSV解析結果は文字列です。先頭ゼロを保つ取込例では転記先を文字列書式にします。数値のCSV化はCStrによる実行環境の表示規則を使います。
- 辞書のキーは文字列として比較します。`1`と`"1"`は同じキー、既定で大文字小文字も同一です。金額などの数値比較はclsAT_RuleでnumericComparisonをTrueにします。
- 吹き出し転記は位置・サイズ・文字・基本フォント・色を対象とします。個別文字の混在書式、グラデーション、吹き出し先端の調整値は対象外です。

## 元ブックからの整理

提供された2つのブックの共通処理を参考に、業務シート・固定設定・自動実行・グローバル状態への依存を外しました。対応表と動作変更は [docs/抽出と変更.md](docs/抽出と変更.md) にあります。元の業務アプリ全体をそのまま実行するものではありません。

## 検証状況

文字コード、手続きの境界、文字列リテラル、重複する名前、依存関係、生成版の一致、公開物の情報チェックを実施しました。**この環境ではExcelでのVBEコンパイル・VBA実行は未実施です。** 完了済みの実行テストとは扱っていません。

Excelで確認する場合は全モジュール・クラスに加え `vbe/tests/mdlAT_SelfTest.bas` を取り込み、コンパイル後 `AT_RunSelfTests` を実行します。列変換16,384列、CSVの引用符・改行、空値・0、条件判定、転記、集計、設定復元などを一時ブックで確認するテストです。詳細は [tests/README.md](tests/README.md) を参照してください。

## 公開とライセンス

公開者はafuro77です。MITライセンスで配布します。許諾条件はLICENSEを参照してください。元ブック、実行履歴、業務データ、保護パスワード、マクロ削除処理は配布物に含めません。
