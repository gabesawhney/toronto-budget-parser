<?php

error_reporting(E_ALL);
set_time_limit(0);

date_default_timezone_set('America/Toronto');

$datafolder = '../citybudget2012/';
$startwithfilesthatbeginwith = 'a';

/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . './Classes/');

/** PHPExcel_IOFactory */
include 'PHPExcel/IOFactory.php';


$allfiles = scandir($datafolder);
foreach ($allfiles as $thisfile) {
	if ( (substr($thisfile, -3) == "xls") OR (substr($thisfile, -4) == "xlsx") ) {
		if ( (substr($thisfile, 0,3) != "zzz") AND (substr($thisfile, 0,1) != "~") ) {
			if ( strcasecmp($startwithfilesthatbeginwith, $thisfile) <= 0 ) {
				$thefiles[] = $thisfile;
			}
		}
	}
}

foreach ($thefiles as $thisfile) {
	$inputFileName = $datafolder . $thisfile;
	//$inputFileName = 'Examples/01simple.xlsx';
	echo 'Loading file '.pathinfo($inputFileName,PATHINFO_BASENAME).": ";

	$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
	$objReader = PHPExcel_IOFactory::createReader($inputFileType);
	$objReader->setReadDataOnly(true);
	$objPHPExcel = $objReader->load($inputFileName);


	$sheetObj = $objPHPExcel->getActiveSheet();
//	$sheetData = $sheetObj->toArray(null,true,true,true);

	//get rid of empty cells
	$maxCell = $sheetObj->getHighestRowAndColumn();
	$sheetData = $sheetObj->rangeToArray('A1:' . $maxCell['column'] . $maxCell['row'],null,true,true,true);

	//get rid of empty cells (part 2)
	foreach($sheetData as $key => &$row) {
		$row = array_filter($row,
							function($cell) {
								if ( !is_null($cell) && ($cell != "") ) return true;
									else return false;
							}
			   );
		if (count($row) == 0) {
			unset($sheetData[$key]);
		}
	}
	unset ($row);

//var_dump($sheetData);

	if ($sheetData[1]["B"] == "") {
		if ($sheetData[1]["A"] == "") {
			$name = "ERROR*******";
		} else {
			$name = $sheetData[1]["A"];
		}
	} else {
		$name = $sheetData[1]["B"];
	}
	//echo $name.PHP_EOL;
	
	//REMOVE BLANK COLUMNS
	$column = 'A';
	$lastRow = $sheetObj->getHighestRow();
	for ($row = 1; $row <= $lastRow; $row++) {
    	$cell = $sheetObj->getCell($column.$row);
		if ($cell == "") {
			// do nothing
		} else {
			echo $cell.PHP_EOL;
		}
	}

var_dump($sheetData[1]);

	unset($sheetData);
}

?>
