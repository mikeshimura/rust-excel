#[cfg(test)]
mod sample{
    use umya_spreadsheet::BorderStyleValues;
    use umya_spreadsheet::*;
    use rustexcel::*;
    #[test]
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
        re.set_cell("B1", "RustExcelTest");
        re.set_cell("E1", "Masanobu Shimura");
        re.set_vertical_aliginment("E1:G1", VerticalAlignmentValues::Bottom);
        re.set_border_style("A1:G1", "tblr", BorderStyleValues::Thick);
        re.set_border_all("A3:G5", "tblr", BorderStyleValues::Thin);
        re.set_border_style("A3:G5", "tblr", BorderStyleValues::Thick);
        re.set_background_color("A3:G3", "dcdcdc");
        re.set_font_style_bold("A3:G3", true);
        re.set_cell("A3", "Date");
        re.set_cell("B3", "String");
        re.set_cell("C3", "Italic");
        re.set_cell("D3", "Bold");
        re.set_cell("E3", "Underline");
        re.set_cell("F3", "Number");
        re.set_cell("G3", "Number");
        re.set_cell_date("A4", 2025, 03, 03, 12, 00, 00);
        re.set_number_format("A4", "yyyy/mm/dd hh:mm".to_string());
        re.set_cell("B4", "RustExcel");
        re.set_cell("C4", "RustExcel");
        re.set_cell("D4", "RustExcel");
        re.set_cell("E4", "RustExcel");
        re.set_font_style_italic("C4", true);
        re.set_font_style_bold("D4", true);
        re.set_font_style_under_line("E4", UnderlineValues::Single);
        re.set_cell_number("F4", 123456.0);
        re.set_number_format("F4", "#,##0.00".to_string());
        re.set_cell_number("G4", 12345.0);
        re.set_number_format("G4", "#,##0".to_string());
        re.save("sample.xlsx");

    }

}