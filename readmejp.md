# RustExcel: Rust での Excel ファイル操作ライブラリ

RustExcel は、Excel (.xlsx) ファイルの作成と操作を簡素化するために設計された Rust ライブラリです。Excel スプレッドシートを操作するための高レベルな `Context` オブジェクトを提供し、一般的なタスクをより簡単に実行できるようにします。内部では、RustExcel は .xlsx ファイル形式の複雑な処理に `umya_spreadsheet` クレートを活用しています。

## 主な機能

-   **コンテキストベースの操作:** RustExcel のコアは `Context` 構造体であり、Excel スプレッドシートを管理するための中央拠点として機能します。
-   **シート管理:**
    -   指定された名前で新しいシートを作成します。
    -   インデックスまたは名前でアクティブなシートを設定します。
    -   変更のためにアクティブなシートにアクセスします。
-   **セル操作:**
    -   セルに文字列、数値、または日付の値を設定します。
    -   より高度な操作のためにセルオブジェクトを取得します。
    -   座標 ("A1" など) または行番号と列番号でセルにアクセスします。
-   **スタイル管理:**
    -   フォント プロパティ (名前、サイズ、太字、斜体、下線、取り消し線) を設定します。
    -   フォントの色を設定します。
    -   罫線のスタイルを設定します。
    -   背景色を設定します。
    -   数値形式を設定します。
- **日付**:
    - セルに日付を設定します。
- **範囲**:
    - 範囲を解析します。

## コア コンポーネント

### `Context`

`Context` 構造体は、RustExcel で Excel ファイルを操作するための主要なインターフェイスです。`umya_spreadsheet` の `Spreadsheet` オブジェクト、アクティブなシート、およびその他のコンテキスト情報を管理します。

**主なメソッド:**

-   `new()`: 新規の空の Excel ファイル コンテキストを作成します。
-   `get_book()`: 基になる `Spreadsheet` オブジェクトへの可変参照を返します。
-   `new_sheet(name: &str)`: 指定された名前で新しいシートを作成し、アクティブなシートとして設定します。
-   `set_sheet_by_index(index: usize)`: インデックス (0 ベース) でアクティブなシートを設定します。
-   `set_sheet_by_name(name: &str)`: 名前でアクティブなシートを設定します。
-   `get_sheet()`: アクティブなシートへの参照を返します。
-   `get_sheet_mut()`: アクティブなシートへの可変参照を返します。
-   `set_row(row: u32)`: 現在の行を設定します。
-   `get_row()`: 現在の行を取得します。
-   `set_cell(cell: &str, value: &str)`: セルの文字列値を設定します (例: "A1")。
- `set_cell_number(cell: &str, value: f64)`: セルに数値を設定します。
- `set_cell_date(cell: &str, year: u32, month: u32, day: u32, hour: u32, minute: u32, second: u32)`: セルに日付を設定します。
-   `set_cell_by_coordinate(row: u32, col: u32, value: &str)`: 行番号と列番号でセルの値を設定します。
-   `get_cell(cell: &str)`: 座標でセルへの可変参照を取得します。
-   `get_cell_by_coordinate(row: u32, col: u32)`: 行番号と列番号でセルへの可変参照を取得します。
-   `get_cell_by_col(col: u32)`: 現在の行の列番号でセルへの可変参照を取得します。
-   `get_style(cell: &str)`: セルのスタイルへの可変参照を取得します。
-   `get_style_by_coordinate(row: u32, col: u32)`: 行番号と列番号でセルのスタイルへの可変参照を取得します。
-   `get_style_by_col(col: u32)`: 現在の行の列番号でセルのスタイルへの可変参照を取得します。
    -   `set_font_size(range: &str, font_size: f64)`: 範囲内のフォント サイズを設定します。
    -   `set_font_style_bold(range: &str, font_style: bool)`: 範囲内のフォントを太字に設定します。
    -  `set_font_style_italic(range: &str, font_style: bool)`: 範囲内のフォントを斜体に設定します。
    -   `set_font_style_under_line(range: &str, value: UnderlineValues)`: 範囲内のフォントに下線を設定します。
    -   `set_font_style_strike(range: &str, value: bool)`: 範囲内のフォントに取り消し線を設定します。
    -   `set_font_name(range: &str, value: &str)`: 範囲内のフォント名を設定します。
    - `set_font_char_set(range: &str, value: FontCharSet)`: フォントの文字セットを設定します。
    - `set_font_color(range: &str, value: &str)`: 範囲内のフォントの色を設定します。
- `set_border_style(range: &str, pos: &str, border_style: BorderStyleValues)`: 範囲内の罫線のスタイルを設定します。
- `set_border_all(range: &str, pos: &str, border_style: BorderStyleValues)`: 範囲内のすべての罫線を設定します。
    -   `set_background_color(range: &str, value: &str)`: 範囲内の背景色を設定します。
    -   `set_number_format(range: &str, value: String)`: 範囲内の数値形式を設定します。

-   `save(path: &str)`: 指定されたパスに Excel ファイルを保存します。

## 使用例

```rust
use rustexcel::; 
use umya_spreadsheet::;
fn main() { // 新しいコンテキストを作成します let mut context = Context::new();
// "Sheet1" という名前の新しいシートを作成します
context.new_sheet("Sheet1");
// セルを設定します
context.set_cell("A1", "こんにちは、世界！");
context.set_cell_number("A2", 123.45);
context.set_cell_by_coordinate(3, 2, "B3 の値");
context.set_cell_date("C4", 2024, 1, 1, 10, 30, 00);
// フォントのスタイルを設定します
context.set_font_size("A1:C4", 20.0);
context.set_font_style_bold("A1:C4", true);
context.set_font_style_italic("A1:C4", true);
context.set_font_style_under_line("A1:C4", UnderlineValues::Single);
context.set_font_style_strike("A1:C4", true);
context.set_font_name("A1:C4", "ＭＳ ゴシック");
context.set_font_char_set("A1:C4", FontCharSet::default());
context.set_font_color("A1:C4", "FF0000");
// 罫線のスタイルを設定します。 context.set_border_style( " B2: C4" , " tblr" , BorderStyleValues: : Thin) ;  context.set_border_all( " B2: C4" , " tblr" , BorderStyleValues: : Thick) ;  // 背景色を設定します context.set_background_ color( " A1: C4" ,  "00FFFF");
// 数値形式を設定します
context.set_number_format("C4", "yyyy/mm/dd hh:mm:ss".to_string());
// ファイルを保存します
context.save("output.xlsx");
}
```
## 依存関係

-   `umya_spreadsheet`: 中核となる Excel ファイル処理ライブラリ。
-   `regex`: セル範囲の解析に使用します。
-   `chrono`: 日付と時刻の管理に使用します。

## インストール方法

`Cargo.toml` ファイルに以下を追加します。

```toml
[dependencies]
rustexcel = "0.1" # 実際のバージョン番号に置き換えてください
umya_spreadsheet = "2.2.3" 
regex = "1" 
chrono = "0.4"
```
## ライセンス

[ここにライセンス情報を追加してください]

## 貢献

[ここに貢献ガイドラインを追加してください]

