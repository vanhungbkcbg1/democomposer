
<?php
require __DIR__ . '/vendor/autoload.php';
$objPHPExcel = new PHPExcel();


// Set document properties
$objPHPExcel->getProperties()->setCreator("Thouhedul islam")
    ->setLastModifiedBy("Thouhedul islam")
    ->setTitle("PHPExcel Tutorial from tisuchi.com")
    ->setSubject("PHPExcel Tutorial from tisuchi.com")
    ->setDescription("This is the tutorial for PHP Excel from tisuchi.com")
    ->setKeywords("office PHPExcel php")
    ->setCategory("Tutorial Result");


// Add Data in your file

$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1', 'Visit ')
    ->setCellValue('B1', 'tisuchi.com')
    ->setCellValue('C1', 'for interesting')
    ->setCellValue('D1', 'tutorail');



$objPHPExcel->getActiveSheet()->setCellValue('A8',"Posted in \n tisuchi.com");
$objPHPExcel->getActiveSheet()->getRowDimension(8)->setRowHeight(-1);
$objPHPExcel->getActiveSheet()->getStyle('A8')->getAlignment()->setWrapText(true);



// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle('tisuchi.com');


// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$objPHPExcel->setActiveSheetIndex(0);


// Save Excel 2007 file

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save(str_replace('.php', '.xls', __FILE__));



?>