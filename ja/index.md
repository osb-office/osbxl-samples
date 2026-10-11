[English](../en/index.md)

# osbxl サンプル集

実行して確認済みの osbxl サンプルです。各ページのコードはそのまま動きます。 コードは **Node.js / Python / Java / PHP / WebAssembly の5言語タブ**で切り替えられ、 タブごとに表示されるインストール手順も言語別です。

**はじめての方は [general-01 概要 — オブジェクトモデルと基本の流れ](general-01-overview.md) から。**

## やりたいことから引く

| やりたいこと | サンプル |
| --- | --- |
| osbxl の全体像・クラスの階層を知りたい | [general-01 概要](general-01-overview.md) |
| 新規作成 / 既存ブックを開く / save / saveAs / close の流れ | [general-01 概要 — 基本の流れ](general-01-overview.md#基本の流れ) |
| App / WorkBook / WorkSheet / Range にはどんなメソッドがあるか | [general-01 概要 — 主要メソッド](general-01-overview.md#主要メソッド) |
| 表を作って値と数式を入れたい | [general-02 値と数式](general-02-values-formulas.md) |
| 行ごとにずれる数式（相対参照）を一括で入れたい | [general-02 値と数式](general-02-values-formulas.md#コード) / [general-06 一括投入](general-06-bulk-input.md) |
| シートを追加・削除・移動・リネームしたい | [general-03 シート操作](general-03-sheets.md) |
| シート名が分からないブックの全シートを列挙したい | [general-03 シート操作](general-03-sheets.md#コード) |
| 日付・時刻・真偽値を入れたい / 表示書式を指定したい | [general-04 セル値の型](general-04-value-types.md) |
| `getValue()` と `getValue(true)` の違い | [general-04 セル値の型](general-04-value-types.md#getvalue-と-getvaluetrue) |
| 行や列を挿入・削除・クリア・コピーしたい | [general-05 行・列・セル範囲](general-05-rows-cols.md) |
| 行の高さ・列の幅・非表示を設定したい | [general-05 行・列・セル範囲](general-05-rows-cols.md#コード-2-row-col-オブジェクト) |
| 数万セルを高速に投入したい | [general-06 一括投入](general-06-bulk-input.md) |
| データを絞り込む・並べ替えたい（オートフィルター / Sort） | [general-07 オートフィルターと並べ替え](general-07-autofilter-sort.md) |
| テーブル（集計行・構造化参照）やコメントを使いたい | [general-08 テーブルとコメント](general-08-tables-comments.md) |
| 条件付き書式（カラースケール・データバー・上位 N）を設定したい | [general-09 条件付き書式](general-09-conditional-format.md) |
| グラフ（縦棒・折れ線）を作成したい | [general-10 チャート](general-10-charts.md) |
| ピボットテーブルで集計したい | [general-11 ピボットテーブル](general-11-pivottables.md) |
| 他ライブラリ（POI/openpyxl/PHP）との詳細機能対比・ベンチマーク | [ベンチマーク比較実測レポート](benchmark-comparison.md) |

## general — 概要・基本操作

| # | タイトル | level | 主な API |
| --- | --- | --- | --- |
| 01 | [概要 — オブジェクトモデルと基本の流れ](general-01-overview.md) | basic | App / WorkBook / WorkSheet / Range |
| 02 | [値と数式を入れて表を作る](general-02-values-formulas.md) | basic | Range / setRangeFormulas |
| 03 | [ブックとワークシートの操作 — 追加・削除・移動・リネーム](general-03-sheets.md) | basic | WorkBook / WorkSheet |
| 04 | [セル値の型と表示書式 — 文字列・数値・日付・時刻・真偽値](general-04-value-types.md) | basic | Range / DateTimeObject |
| 05 | [行・列・セル範囲の操作 — 挿入・削除・クリア・コピー](general-05-rows-cols.md) | basic | WorkSheet / Row / Col |
| 06 | [大量データの一括高速投入 — setRangeData / setRangeFormulas](general-06-bulk-input.md) | intermediate | WorkSheet (setRangeData / setRangeFormulas) |
| 07 | [オートフィルターと並べ替え — 条件フィルター・ソート](general-07-autofilter-sort.md) | intermediate | AutoFilter / Sort |
| 08 | [テーブルとコメント — ListObjects・集計行・構造化参照](general-08-tables-comments.md) | intermediate | ListObjects / Table / Comments |
| 09 | [条件付き書式 — セル値ルール・カラースケール・データバー](general-09-conditional-format.md) | intermediate | FormatConditions / ColorScaleObject / DataBarObject |
| 10 | [チャートの基本 — 縦棒・折れ線・タイトル・軸・凡例](general-10-charts.md) | intermediate | Chart / chart.\* / enums.XlChartType |
| 11 | [ピボットテーブル — 作成・集計・集計方法の変更](general-11-pivottables.md) | advanced | PivotTable / PivotField |

## 予定（未作成）

### general

ブックの属性（`setCompanyName` / `setCreateAuthor` 他） / 名前付き範囲（`getNames`） / カスタムスタイル（`addCustomStyle` / `addCustomTableStyle`） / 外部ブック連携（`addExternalWorkBook` / `updateExternalWorkBook`） / パスワード付きブック（`openPasswordWorkBook` / `save(password)`） / アウトラインの詳細

### format

フォント・色 / 罫線 / 塗りつぶし・グラデーション / 配置と折り返し / 表示形式（数値・日付・通貨） / 組み込みスタイル・カスタムスタイル （条件付き書式は [general-09](general-09-conditional-format.md) 参照）

### data

セル結合 / ハイパーリンク （テーブル・並べ替え・オートフィルター・コメントは [general-07](general-07-autofilter-sort.md) / [general-08](general-08-tables-comments.md) 参照）

### formula

数式の基本と配列数式 / 財務関数（ACCINT 他） / 検索関数（XLOOKUP） / 予測（FORECAST.ETS） / WEBSERVICE + FILTERXML

### pivot

集計とグループ化（作成・集計の基本は [general-11](general-11-pivottables.md) 参照）

### chart

グラフの種類 / データラベル / 書式（色・線・影） （基本作成・軸・凡例は [general-10](general-10-charts.md) 参照）

### advanced

ページ設定と印刷 / 保護 / ウィンドウ・枠線 / Shapes
