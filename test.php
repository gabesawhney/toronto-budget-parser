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
	echo pathinfo($inputFileName,PATHINFO_BASENAME).": ";

	$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
	$objReader = PHPExcel_IOFactory::createReader($inputFileType);
	$objReader->setReadDataOnly(true);
	$objPHPExcel = $objReader->load($inputFileName);


	$sheetObj = $objPHPExcel->getActiveSheet();
//	$sheetData = $sheetObj->toArray(null,true,true,true);

	//get rid of empty cells
	$maxCell = $sheetObj->getHighestRowAndColumn();
	$sheetData = $sheetObj->rangeToArray('A1:' . $maxCell['column'] . $maxCell['row'],null,true,true,false);

/* forget about this now that we're looking for keywords
 
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
*/

//	$search = "Salaries and Benefits";
	$search = "alarie";
	foreach($sheetData as $colnum => $row) {
		foreach($row as $rownum => $cell) {
//			if ($cell == $search) {
//			if ( strpos($cell,$search) > 0 ) {
			if ( preg_match('/alarie/', $cell) ) {
				print $cell.PHP_EOL;
			} else {
				print "NO MATCH: ".$cell.PHP_EOL;
			}
		}
	}

	unset($sheetData);
}



?>
