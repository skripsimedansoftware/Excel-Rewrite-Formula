<?php
require('vendor/autoload.php');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$reader = new PhpOffice\PhpSpreadsheet\Reader\Xlsx;
$load_file = $reader->load('DATA-TRANSAKSI-PENJUALAN.xlsx');

for ($i = 0; $i < 1000; $i++)
{
	$load_file->getActiveSheet()->setCellValue('F'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B3, C'.($i+13).')))');
	$load_file->getActiveSheet()->setCellValue('G'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B4, C'.($i+13).')))');
	$load_file->getActiveSheet()->setCellValue('H'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B5, C'.($i+13).')))');
	$load_file->getActiveSheet()->setCellValue('I'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B6, C'.($i+13).')))');
	$load_file->getActiveSheet()->setCellValue('J'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B7, C'.($i+13).')))');
	$load_file->getActiveSheet()->setCellValue('K'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B8, C'.($i+13).')))');
	$load_file->getActiveSheet()->setCellValue('L'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B9, C'.($i+13).')))');
	$load_file->getActiveSheet()->setCellValue('M'.($i+13), '=IF(ISNUMBER(SEARCH(B2, C'.($i+13).')), ISNUMBER(SEARCH(B10, C'.($i+13).')))');
}

// Count total of true
$load_file->getActiveSheet()->setCellValue('F1013', '=COUNTIF(F13 : F1012, TRUE)');
$load_file->getActiveSheet()->setCellValue('G1013', '=COUNTIF(G13 : G1012, TRUE)');
$load_file->getActiveSheet()->setCellValue('H1013', '=COUNTIF(H13 : H1012, TRUE)');
$load_file->getActiveSheet()->setCellValue('I1013', '=COUNTIF(I13 : I1012, TRUE)');
$load_file->getActiveSheet()->setCellValue('J1013', '=COUNTIF(J13 : J1012, TRUE)');
$load_file->getActiveSheet()->setCellValue('K1013', '=COUNTIF(K13 : K1012, TRUE)');
$load_file->getActiveSheet()->setCellValue('L1013', '=COUNTIF(L13 : L1012, TRUE)');
$load_file->getActiveSheet()->setCellValue('M1013', '=COUNTIF(M13 : M1012, TRUE)');

$writer = new Xlsx($load_file);
$writer->save('export.xlsx');
?>
