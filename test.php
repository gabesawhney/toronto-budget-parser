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
	echo pathinfo($inputFileName,PATHINFO_BASENAME).": ".PHP_EOL;

	$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
	$objReader = PHPExcel_IOFactory::createReader($inputFileType);
	$objReader->setReadDataOnly(true);
	$objPHPExcel = $objReader->load($inputFileName);


	$sheetObj = $objPHPExcel->getActiveSheet();
//	$sheetData = $sheetObj->toArray(null,true,true,true);

	//get rid of empty cells
	$maxCell = $sheetObj->getHighestRowAndColumn();
	$sheetData = $sheetObj->rangeToArray('A1:' . $maxCell['column'] . $maxCell['row'],null,false,false,false);

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
	unset($row);

foreach($sheetData as $key => $row) {
//print count($row);
	print sprintf("%2d ",count($row));
}
print PHP_EOL;

	unset($sheetData);
}



?>
