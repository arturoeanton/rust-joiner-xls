mod simple_excel_writer;

use simple_excel_writer::*;

use calamine::{open_workbook, DataType, Error, Reader, Xlsx};
use std::collections::hash_map::HashMap;

use clap::{App, Arg};

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
    format_date: &str
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
        let mut only_fields_names:Vec<String> = Vec::new(); 
        let mut data_fix_map: HashMap<String,String> = HashMap::new();
        for field in fields.iter() {
            let params: Vec<&str> = field.split(":as:").collect();
            let param1 = params[0].to_string();
            let mut key_name = param1.to_string();
            let mut alias_name = key_name.to_string();
            if params.len() == 2 { // tenes alias
                alias_name = params[1].to_string();
            }
            let subparams:Vec<&str>  = param1.split("=").collect();
            if subparams.len() == 2 {
                key_name = subparams[0].to_string().trim().to_string();
                let value =  subparams[1].to_string().trim().to_string();
                data_fix_map.insert(key_name.to_string(), value);
                if params.len() != 2 {
                    alias_name = key_name.to_string();
                }
            }
            println!("{}",alias_name);
            println!("{}",key_name);
            println!("----");

            row.add_cell(alias_name, CellStyle::BoldCenter);
            only_fields_names.push(key_name);
        }
        let mut result = sw.append_row(row);

        for page_row in page.iter() {
            let mut row_writer = Row::new();
            for field in only_fields_names.iter() {
                let key = String::from(field.to_string());
                let data1 = page_row.get(key.trim());
                match data1 {
                    Some(dt) => match dt {
                        DataType::Int(value) => {
                            let cell = CellValue::Number(*value as f64);
                            row_writer.add_cell(cell, CellStyle::BoldLeft);
                        }
                        DataType::Float(value) => {
                            let cell = CellValue::Number(*value);
                            row_writer.add_cell(cell, CellStyle::BoldLeft);
                        }
                        DataType::Bool(value) => {
                            let cell = CellValue::Bool(*value);
                            row_writer.add_cell(cell, CellStyle::BoldLeft);
                        }
                        DataType::DateTime(value) => {
                            let unix_days = value - 25569.;
                            let unix_secs = unix_days * 86400.;
                            let secs = unix_secs.trunc() as i64;
                            let nsecs = (unix_secs.fract().abs() * 1e9) as u32;
                            let ch = chrono::NaiveDateTime::from_timestamp_opt(secs, nsecs);

                            let value = format!("{}", ch.unwrap().format(format_date));
                            let cell = CellValue::String(value);
                            row_writer.add_cell(cell, CellStyle::BoldLeft);
                        }
                        DataType::String(value) => {
                            let cell = CellValue::String(value.to_string());
                            row_writer.add_cell(cell, CellStyle::Left)
                        }
                        DataType::Empty =>{
                            row_writer.add_empty_cells(1);
                        }
                        _ => {
                            row_writer.add_empty_cells(1);
                        }
                    },
                    None => {
                            let value_fix = data_fix_map.get(field);
                            println!("{}",field);
                            let mut v = value_fix.unwrap().to_string();
                            match (v.chars().nth(0), v.chars().rev().nth(0)) {
                                (Some('\''), Some('\'')) => {
                                    v.pop();
                                    v = v[1..].to_string()
                                }
                                _ => {}
                            }
                            row_writer.add_cell(v, CellStyle::Left);
                        
                    }
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
    let matches = App::new("rust-joiner-excel")
        .version("1.0")
        .about("Rust Joiner Excel")
        .arg(
            Arg::with_name("file1")
                .short("1")
                .long("file1")
                .value_name("FILE")
                .help("File 1")
                .required(true)
                .takes_value(true),
        )
        .arg(
            Arg::with_name("file2")
                .short("2")
                .long("file2")
                .value_name("FILE")
                .help("File 2"),
        )
        .arg(
            Arg::with_name("file_out")
                .short("o")
                .long("file_out")
                .value_name("FILE")
                .help("File Out")
                .required(true)
                .takes_value(true),
        )
        .arg(
            Arg::with_name("sheet1")
                .short("a")
                .long("sheet1")
                .value_name("Sheet1")
                .help("Sheet 1")
                .required(true),
        )
        .arg(
            Arg::with_name("sheet2")
                .short("b")
                .long("sheet2")
                .value_name("Sheet2")
                .help("Sheet 2"),
        )
        .arg(
            Arg::with_name("sheet_out")
                .short("s")
                .long("sheet_out")
                .value_name("SheetOut")
                .help("Sheet Out"),
        )
        .arg(
            Arg::with_name("field_match1")
                .short("x")
                .long("field_match1")
                .value_name("Field Match 1")
                .help("Field Match 1")
                .required(true),
        )
        .arg(
            Arg::with_name("field_match2")
                .short("y")
                .long("field_match2")
                .value_name("Field Match 2")
                .help("Field Match 2")
                .required(true),
        )
        .arg(
            Arg::with_name("distinct")
                .short("d")
                .long("distinct")
                .help("Distinct"),
        )
        .arg(
            Arg::with_name("fields_output")
                .short("O")
                .value_name("Fields Output")
                .long("fields_output")
                .required(true)
                .help("Fields Output"),
        )
        .arg(
            Arg::with_name("format_date")
                .value_name("Format Date")
                .long("%m/%d/%Y")
                .help("Fields Output"),
        )
        .get_matches();

    let format_date = matches.value_of("format_date").unwrap_or("%m/%d/%Y") ;//"%Y-%m-%d %H:%M:%S";
    let field_output = matches.value_of("fields_output").unwrap().to_string();
    let field_match1 = matches.value_of("field_match1").unwrap().to_string();
    let field_match2 = matches.value_of("field_match2").unwrap().to_string();
    let distinct = matches.is_present("distinct");
    let name_file1 = matches.value_of("file1").unwrap();
    let name_file2 = matches.value_of("file2").unwrap_or(name_file1);
    let name_file_out = matches.value_of("file_out").unwrap();
    let sheet_name1 = matches.value_of("sheet1").unwrap_or("Sheet1");
    let sheet_name2 = matches.value_of("sheet2").unwrap_or(sheet_name1);
    let sheet_name_out = matches.value_of("sheet_out").unwrap_or("Sheet1");

    let page1 = reader_xlsx(&name_file1, &sheet_name1).unwrap();
    let page2 = reader_xlsx(&name_file2, &sheet_name2).unwrap();

    let page_out = merge_pages(&page1, &page2, &field_match1, &field_match2, &distinct).unwrap();

    let fieds: Vec<&str> = field_output.split(",").collect();
    let _ = create_new_excel(&name_file_out, &sheet_name_out, &fieds, &page_out, format_date);
}
/*
cargo run -- \
    --file1 "./test_files/test_dup.xlsx"  --file_out "./out_files/test_dup1.xlsx"  --sheet1 "Vista Qlik"  --sheet2 "Spool (SISE)"  \
    --field_match1 numpol \
    --field_match2 Poliza \
    --fields_output "Poliza, numpol, Chasis,desmotor, producto='pepe pe', codepais=ar, Zona Riesgo"

*/
