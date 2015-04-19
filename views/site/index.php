<?php
/* @var $this yii\web\View */
$this->title = 'My Yii Application';

//var_dump($objPHPExcel);
//Get worksheet dimensions
$sheet = $objPHPExcel->getSheet(3);
$highestRow = $sheet->getHighestRow();
$highestColumn = $sheet->getHighestColumn();

//  Loop through each row of the worksheet in turn
for ($row = 1; $row <= $highestRow; $row++) {
    //  Read a row of data into an array
    $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row,
        NULL, TRUE, FALSE);
    foreach($rowData[0] as $k=>$v)
        echo "Row: ".$row."- Col: ".($k+1)." = ".$v."<br />";
}
?>

