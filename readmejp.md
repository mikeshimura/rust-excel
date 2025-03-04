# RustExcel: Rust での Excel ファイル操作ライブラリ

RustExcel は、Excel (.xlsx) ファイルの作成と操作を簡素化するために設計された Rust ライブラリです。Excel スプレッドシートを操作するための高レベルな `RustExcel` オブジェクトを提供し、一般的なタスクをより簡単に実行できるようにします。内部では、RustExcel は .xlsx ファイル形式の複雑な処理に `umya_spreadsheet` クレートを活用しています。
![Sample Xlsx](https://raw.githubusercontent.com/mikeshimura/rust-excel/refs/heads/master/sample.png
)

## サンプル　ソースコード
```rust
fn sample() {
    let mut re = RustExcel::new();
    let mut book: &mut Spreadsheet = re.get_book();
    re.new_sheet("Sample");
    re.set_column_width(1,1,17.5);
    re.set_column_width(2,7,10.0);
    re.set_row_height(1,1,33.8);
    re.set_font_color("A1:G1", "0067C0");
    re.set_background_color("A1:G1", "FFFF00");
    re.set_font_size("A1:D1", 20.0);
    re.set_font_size("E1:G1", 14.0);
    re.set_cell_string("B1", "RustExcelTest");
    re.set_cell_string("E1", "Masanobu Shimura");
    re.set_vertical_aliginment("E1:G1", VerticalAlignmentValues::Bottom);
    re.set_border_style("A1:G1", "tblr", BorderStyleValues::Thick);
    re.set_border_all("A3:G5", "tblr", BorderStyleValues::Thin);
    re.set_border_style("A3:G5", "tblr", BorderStyleValues::Thick);
    re.set_background_color("A3:G3", "dcdcdc");
    re.set_font_style_bold("A3:G3", true);
    re.set_row(3);
    re.set_cell_string_by_col_number(1, "Date");
    re.set_cell_string_by_col_number(2, "String");
    re.set_cell_string_by_col_number(3, "Italic");
    re.set_cell_string_by_col_number(4, "Bold");
    re.set_cell_string_by_col_number(5, "Underline");
    re.set_cell_string_by_col_number(6, "Number");
    re.set_cell_string_by_col_number(7, "Number");
    re.set_cell_date("A4", 2025, 03, 03, 12, 00, 00);
    re.set_number_format("A4", "yyyy/mm/dd hh:mm".to_string());
    re.set_cell_string_by_coordinate(4,2, "RustExcel");
    re.set_cell_string_by_coordinate(4,3, "RustExcel");
    re.set_cell_string_by_coordinate(4,4, "RustExcel");
    re.set_cell_string_by_coordinate(4,5, "RustExcel");
    re.set_font_style_italic("C4", true);
    re.set_font_style_bold("D4", true);
    re.set_font_style_under_line("E4", UnderlineValues::Single);
    re.set_cell_number("F4", 123456.0);
    re.set_number_format("F4", "#,##0.00".to_string());
    re.set_cell_number("G4", 12345.0);
    re.set_number_format("G4", "#,##0".to_string());
    re.save("sample.xlsx");

    }
```
## 主な機能

-   **RustExcel構造体ベースの操作:** RustExcel のコアは `RustExcel` 構造体であり、Excel スプレッドシートを管理するための中央拠点として機能します。
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
    -   文字列の水平・垂直配置を設定します。
    -   罫線のスタイルを設定します。
    -   背景色を設定します。
    -   数値形式を設定します。
- **日付**:
    - セルに日付を設定します。
- **範囲**:
    - 範囲を解析します。

## コア コンポーネント

### `RustExcel`

`RustExcel` 構造体は、RustExcel で Excel ファイルを操作するための主要なインターフェイスです。`umya_spreadsheet` の `Spreadsheet` オブジェクト、アクティブなシート、およびその他のRustExcel構造体情報を管理します。

**主なメソッド:**

-   `new()`: 新規の空の Excel ファイル RustExcel構造体を作成します。
-   `read(path)`: Excel ファイル を読み込み RustExcel構造体を作成します。
-   `get_book()`: 基になる `Spreadsheet` オブジェクトへの可変参照を返します。
-   `new_sheet(name: &str)`: 指定された名前で新しいシートを作成し、アクティブなシートとして設定します。
-   `set_sheet_by_index(index: usize)`: インデックス (0 ベース) でアクティブなシートを設定します。
-   `set_sheet_by_name(name: &str)`: 名前でアクティブなシートを設定します。
-   `get_sheet()`: アクティブなシートへの参照を返します。
-   `get_sheet_mut()`: アクティブなシートへの可変参照を返します。
-   `set_row(row: u32)`: 現在の行を設定します。
-   `get_row()`: 現在の行を取得します。
-   `set_row_height(row_from: u32,row_to:u32, height: f64) )`: 行の高さを設定します。
-   `set_column_width(col_from: u32,col_to:u32, width: f64)`: 列の幅を設定します。
-   `set_cell_string(cell: &str, value: &str)`: セルの文字列値を設定します (例: "A1")。
- `set_cell_number(cell: &str, value: f64)`: セルに数値を設定します。
- `set_cell_date(cell: &str, year: u32, month: u32, day: u32, hour: u32, minute: u32, second: u32)`: セルに日付を設定します。
-   `set_cell_string_by_coordinate(row: u32, col: u32, value: &str)`: 行番号と列番号でセルの文字列値を設定します。
-   `set_cell_number_by_coordinate(row: i32, col: i32, value: f64)`: 行番号と列番号でセルの数値を設定します。
-   `set_cell_date_by_coordinate(&mut self,row: i32,col: i32,year: u32,month: u32,day: u32,hour: u32,minute: u32,second: u32)`: 行番号と列番号でセルに日付を設定します
-   `get_cell(cell: &str)`: 座標でセルへの可変参照を取得します。
-   `get_cell_by_coordinate(row: u32, col: u32)`: 行番号と列番号でセルへの可変参照を取得します。
-   `get_cell_by_col(col: u32)`: 現在の行の列番号でセルへの可変参照を取得します。
-   `get_style(cell: &str)`: セルのスタイルへの可変参照を取得します。
-   `get_style_by_coordinate(row: u32, col: u32)`: 行番号と列番号でセルのスタイルへの可変参照を取得します。
-   `get_style_by_col(col: u32)`: 現在の行の列番号でセルのスタイルへの可変参照を取得します。
-   `set_font_size(range: &str, font_size: f64)`: 範囲内のフォント サイズを設定します。
-   `set_font_style_bold(range: &str, font_style: bool)`: 範囲内のフォントを太字に設定します。
-   `set_font_style_italic(range: &str, font_style: bool)`: 範囲内のフォントを斜体に設定します。
-   `set_font_style_under_line(range: &str, value: UnderlineValues)`: 範囲内のフォントに下線を設定します。
-   `set_font_style_strike(range: &str, value: bool)`: 範囲内のフォントに取り消し線を設定します。
-   `set_font_name(range: &str, value: &str)`: 範囲内のフォント名を設定します。
-   `set_font_char_set(range: &str, value: FontCharSet)`: フォントの文字セットを設定します。
-   `set_font_color(range: &str, value: &str)`: 範囲内のフォントの色を設定します。
-   `set_border_style(range: &str, pos: &str, border_style: BorderStyleValues)`: 範囲内の罫線のスタイルを設定します。
-   `set_border_all(range: &str, pos: &str, border_style: BorderStyleValues)`: 範囲内のすべての罫線を設定します。
-   `set_background_color(range: &str, value: &str)`: 範囲内の背景色を設定します。
-   `set_number_format(range: &str, value: String)`: 範囲内の数値形式を設定します。
-   `set_horizontal_aliginment(range: &str, value: HorizontalAlignmentValues)`:水平方向の文字列配置を設定します。
-   `set_vertical_aliginment(range: &str, value: VerticalAlignmentValues)`:垂直方向の文字列配置を設定します。
-   `save(path: &str)`: 指定されたパスに Excel ファイルを保存します。

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

MIT

## 貢献

Pull Request はいつでも歓迎します。大きな変更を行う前に、提案をご検討ください。

