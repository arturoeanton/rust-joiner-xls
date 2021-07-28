mod simple_excel_writer ;

use simple_excel_writer::*;

use calamine::{open_workbook, DataType, Error, Reader, Xlsx};
use std::collections::hash_map::HashMap;

fn reader_xlsx(path: &str, sheet1: &str) -> Result<Vec<HashMap<String, DataType>>, Error> {
    let mut page: Vec<HashMap<String, DataType>> = Vec::new();
    let mut workbook: Xlsx<_> = open_workbook(path)?;
    let range = workbook
        .worksheet_range(sheet1)
        .ok_or(Error::Msg("Cannot find sheet1"))??;

    let mut fields: Vec<&str> = Vec::new();
    for row in range.rows() {
        if fields.len() == 0 {
            for item in row.iter() {
                fields.push(item.get_string().unwrap_or_default());
            }
        }
        let mut i = 0;
        let mut vrow: HashMap<String, DataType> = HashMap::new();
        for item in row.iter() {
            let key = fields.get(i).unwrap().to_string();
            vrow.insert(key, item.to_owned());
            i += 1;
        }
        page.push(vrow);
    }
    Ok(page)
}

// function merge_pages for merge pages
fn merge_pages(
    page1: &Vec<HashMap<String, DataType>>,
    page2: &Vec<HashMap<String, DataType>>,
    field_match1: &String,
    field_match2: &String,
    distinct: &bool,
) -> Result<Vec<HashMap<String, DataType>>, Error> {
    let mut page: Vec<HashMap<String, DataType>> = Vec::new();

    for row1 in page1.iter() {
        let value_match1 = row1.get(field_match1);
        if value_match1 == None {
            continue;
        }
        let value_match1 = value_match1.unwrap();
        for row2 in page2.iter() {
            let value_match2 = row2.get(field_match2);
            if value_match2 == None {
                continue;
            }
            let value_match2 = value_match2.unwrap();
            if value_match1 == value_match2 {
                let mut new_row: HashMap<String, DataType> = HashMap::new();
                for (key, item) in row1.iter() {
                    new_row.insert(key.to_string(), item.to_owned());
                }
                for (key, item) in row2.iter() {
                    new_row.insert(key.to_string(), item.to_owned());
                }
                page.push(new_row);
                if distinct == &true {
                    break;
                }
            }
        }
    }

    return Ok(page);
}

fn create_new_excel(
    path: &str,
    sheet1: &str,
    fields: &Vec<&str>,
    page: &Vec<HashMap<String, DataType>>,
) -> Result<(), Error> {
    let mut wb = Workbook::create(path);
    let mut sheet = wb.create_sheet(sheet1);

    for field in fields.iter() {
        let w = field.len() as f32 * 3.0;
        sheet.add_column(Column { width: w });
    }

    wb.write_sheet(&mut sheet, |sheet_writer| {
        let sw = sheet_writer;

        let mut row = Row::new();
        for field in fields.iter() {
            row.add_cell(field.to_string(), CellStyle{index:5});
        }
        let mut result = sw.append_row(row);

        for page_row in page.iter() {
            let mut row_writer = Row::new();
            for field in fields.iter() {
                let key = String::from(field.to_string());
                let data = page_row.get(key.trim());

                match Some(data) {
                    Some(data) => {
                        let dt = data.unwrap();
                        if dt.is_int() {
                            let v = dt.get_int().unwrap_or_default();
                            let cell = CellValue::Number(v as f64);
                            row_writer.add_cell(cell, CellStyle{index:3});
                        } else if dt.is_float() {
                            let v = dt.get_float().unwrap_or_default();
                            let cell = CellValue::Number(v);
                            row_writer.add_cell(cell,CellStyle{index:3});
                        } else if dt.is_bool() {
                            let v = dt.get_bool().unwrap_or_default();
                            let cell = CellValue::Bool(v);
                            row_writer.add_cell(cell,CellStyle{index:3});
                        } else if dt.is_empty() {
                            row_writer.add_empty_cells(1);
                        } else {
                            row_writer.add_cell(dt.get_string().unwrap_or_default(),CellStyle{index:1});
                        }
                    }
                    None => {}
                }
            }
            result = sw.append_row(row_writer);
        }
        result
    })
    .expect("write excel error!");

    wb.close().expect("close excel error!");
    Ok(())
}

#[allow(unused_variables)]
fn main() {
    let field_output = "Poliza, numpol, Chasis,desmotor, Zona Riesgo";
    let field_match1 = "numpol".to_string();
    let field_match2 = "Poliza".to_string();
    let distinct = true;
    let name_file1 = "./test_files/test_dup.xlsx".to_string();
    let name_file2 = "./test_files/test_dup.xlsx".to_string();
    let name_file_out = "./out_files/test_dup.xlsx".to_string();

    let sheet_name1 = "Vista Qlik".to_string();
    let sheet_name2 = "Spool (SISE)".to_string();
    let sheet_name_out = "Sheet1".to_string();

    let page1 = reader_xlsx(&name_file1, &sheet_name1).unwrap();
    let page2 = reader_xlsx(&name_file2, &sheet_name2).unwrap();

    let page_out = merge_pages(&page1, &page2, &field_match1, &field_match2, &distinct).unwrap();

    let fieds: Vec<&str> = field_output.split(",").collect();
    let _ = create_new_excel(&name_file_out, &sheet_name_out, &fieds, &page_out);
}
