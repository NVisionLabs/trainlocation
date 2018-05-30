<?php 
$connect = mysqli_connect("localhost", "root", "", "new", "3308");
$connect->set_charset("utf8");
include("PHPExcel/IOFactory.php");
$html = "<table border='1'>";
$objPHPExcel = PHPExcel_IOFactory::load('index.xls');

foreach ($objPHPExcel->getWorksheetIterator() as $worksheet)
{
	$highestRow = $worksheet->getHighestRow();
	for ($row = 2; $row <=$highestRow; $row++)
	{
		$html.="<tr>";
		$numberCol = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(1, $row)->getValue());
		$operationCol = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(3, $row)->getValue());
		$dateCol = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(4, $row)->getValue());
		$departureCol = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(6, $row)->getValue());
		$arrivalCol = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(8, $row)->getValue());
		$currentStateCol = mysqli_real_escape_string($connect, $worksheet->getCellByColumnAndRow(9, $row)->getValue());
		$sql = mysqli_query($connect,"INSERT INTO trains (number, operation, departure, arrival, currentState, date) 
                        VALUES ('".$numberCol."','".$operationCol."','".$departureCol."','".$arrivalCol."','".$currentStateCol."','".$dateCol."')");
		
		$html.="<td>".$numberCol."</td>";
		$html.="<td>".$operationCol."</td>";
		$html.="<td>".$dateCol."</td>";
		$html.="<td>".$departureCol."</td>";
		$html.="<td>".$arrivalCol."</td>";
		$html.="<td>".$currentStateCol."</td>";
		$html.="</tr>";
	}
}
    if ($sql) {
        echo "<p>Данные успешно добавлены в таблицу.</p>";
    } else {
        echo "<p>Произошла ошибка.</p>";
		echo("Error description: " . mysqli_error($connect));
    }
$html.="</table>";
echo $html;

?>