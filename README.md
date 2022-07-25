# Rust xlsx joiner

Xlsx joiner is a rust application to join two different xlsx sheets by a particular field. 

Those sheets can be in separate files or in one file. 

## Installation

 * Xlsx joiner requires Rust to run. 

## Usage


 * Help command

```rust
cargo run -- -help  

USAGE:
    rust-joiner-xls [FLAGS] [OPTIONS] --field_match1 <Field Match 1> --field_match2 <Field Match 2> --fields_output <Fields Output> --file1 <FILE> --file_out <FILE> --sheet1 <Sheet1>

FLAGS:
    -d, --distinct    Distinct
    -h, --help        Prints help information
    -V, --version     Prints version information

OPTIONS:
    -x, --field_match1 <Field Match 1>     Field Match 1
    -y, --field_match2 <Field Match 2>     Field Match 2
    -O, --fields_output <Fields Output>    Fields Output
    -1, --file1 <FILE>                     File 1
    -2, --file2 <FILE>                     File 2
    -o, --file_out <FILE>                  File Out
    -a, --sheet1 <Sheet1>                  Sheet 1
    -b, --sheet2 <Sheet2>                  Sheet 2
    -s, --sheet_out <SheetOut>             Sheet Out


```

 * Example

```rust
cargo run -- \
    --file1 "./test_files/test_dup.xlsx"  --file_out "./out_files/test_dup1.xlsx"  --sheet1 "Vista Qlik"  --sheet2 "Spool (SISE)"  \
    --field_match1 numpol \
    --field_match2 Poliza \
    --fields_output "Poliza, Chasis, Zona Riesgo" 
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.


## License

Apache 2.0 License
