# rust-joiner-xls


```
cargo run -- \
    --file1 "./test_files/test_dup.xlsx"  --file_out "./out_files/test_dup1.xlsx"  --sheet1 "Vista Qlik"  --sheet2 "Spool (SISE)"  \
    --field_match1 numpol \
    --field_match2 Poliza \
    --fields_output "Poliza, numpol, Chasis,desmotor, producto='pepe pe', codepais=ar, Zona Riesgo"   
```
