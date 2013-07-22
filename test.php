<?php

error_reporting(E_ALL);
set_time_limit(0);

date_default_timezone_set('America/Toronto');

$datafolder = '../citybudget2012/';
$startwithfilesthatbeginwith = 't';

/** Include path **/
set_include_path(get_include_path() . PATH_SEPARATOR . './Classes/');

/** PHPExcel_IOFactory */
include 'PHPExcel/IOFactory.php';

$columnloop[0] = "2009 Actual";
$columnloop[1] = "2010 Actual";
$columnloop[2] = "2011 Budget";
$columnloop[3] = "2011 Projected";
$columnloop[4] = "2012 Budget";
$columnloop[5] = "2013 Outlook";
$columnloop[6] = "2014 Outlook";

$expensesloop[] = "Contributions to Capital";
$expensesloop[] = "Contributions to CCRF Naming Rights";
$expensesloop[] = "Contributions to Reserve/Res Funds";
$expensesloop[] = "Contributions to Reserve Funds";
$expensesloop[] = "Salaries and Benefits";
$expensesloop[] = "Cost of Sales";
$expensesloop[] = "Equipment";
$expensesloop[] = "Interdivisional Charges";
$expensesloop[] = "Materials and Supplies";
$expensesloop[] = "Other Expenditures";
$expensesloop[] = "Services & Rents";

$expenses['salaryrow'] = "Salaries and Benefits";
$expenses['equipmentrow'] = "Equipment";
$expenses['materialsrow'] = "Materials and Supplies";
$expenses['salesrow'] = "Cost of Sales";
$expenses['interdivchargesrow'] = "Interdivisional Charges";
$expenses['capitalrow'] = "Contributions to Capital";
$expenses['ccrfrow'] = "Contributions to CCRF Naming Rights";
$expenses['toreserverow'] = "Contributions to Reserve/Res Funds";
$expenses['toreservefundrow'] = "Contributions to Reserve Funds";
$expenses['otherrow'] = "Other Expenditures";
$expenses['servicesrow'] = "Services & Rents";

$revenuesloop[] = "Interdivisional Recoveries";
$revenuesloop[] = "Provincial Subsidies";
$revenuesloop[] = "Federal Subsidies";
$revenuesloop[] = "Other Subsidies";
$revenuesloop[] = "User Fees & Donations";
$revenuesloop[] = "Transfers from Capital Fund";
$revenuesloop[] = "Contribution from Reserve/Res Funds";
$revenuesloop[] = "Contributions from Reserve Funds";
$revenuesloop[] = "Sundry Revenues";

$revenues['interdivrecorow'] = "Interdivisional Recoveries";
$revenues['provsubrow'] = "Provincial Subsidies";
$revenues['fedsubrow'] = "Federal Subsidies";
$revenues['othersubrow'] = "Other Subsidies";
$revenues['userfeesrow'] = "User Fees & Donations";
$revenues['transcaprow'] = "Transfers from Capital Fund";
$revenues['fromreserverow'] = "Contribution from Reserve/Res Funds";
$revenues['fromreservefundrow'] = "Contributions from Reserve Funds";
$revenues['sundryrevrow'] = "Sundry Revenues";

$outputrevenuesfilename = "revenues.csv";
$outputexpensesfilename = "expenses.csv";
$outputpositionsfilename = "positions.csv";
$outputchartfilename = "chart.csv";
$outputjsonfilename = "budget2012.json";

$outputaslist = 1docs.google.com/feeds/download/spreadsheets/Export?key<FILE_ID>&exportFormat=csv&gid=0;
$outputaschart = 0;
$outputasjson = 0;

if ($outputaslist) {
	unlink($outputrevenuesfilename);
	unlink($outputexpensesfilename);
	unlink($outputpositionsfilename);
} 
if ($outputaschart) {
	unlink($outputchartfilename);
}
if ($outputasjson) {
	unlink($outputjsonfilename);
	$jsoncount = 0;
}

///////////////////////////////////////////////////
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
	//echo pathinfo($inputFileName,PATHINFO_BASENAME).": ";

	$inputFileType = PHPExcel_IOFactory::identify($inputFileName);
	$objReader = PHPExcel_IOFactory::createReader($inputFileType);
	$objReader->setReadDataOnly(true);
	$objPHPExcel = $objReader->load($inputFileName);


	//identify the correct worksheet
	$foundit = null;
	unset($worksheetnames);
	$worksheetarray = $objPHPExcel->getAllSheets();
	if (count($worksheetarray) == 1) {
		$foundit = $worksheetarray[0]->getTitle();
	} else {
		foreach($worksheetarray as $sheet) {
			if ($sheet->getTitle() == "Appendix 2 - Budget by Category") { $foundit = $sheet->getTitle(); }
			elseif ($sheet->getTitle() == "City Mgr.") { $foundit = $sheet->getTitle(); }
			elseif ($sheet->getTitle() == "City Clerk's") { $foundit = $sheet->getTitle(); }
			elseif ($sheet->getTitle() == "City Council") { $foundit = $sheet->getTitle(); }
			elseif ($sheet->getTitle() == "Mayor") { $foundit = $sheet->getTitle(); }
			elseif ($sheet->getTitle() == " Budget by Category") { $foundit = $sheet->getTitle(); }
			elseif ($sheet->getTitle() == "Sheet1") { $foundit = $sheet->getTitle(); }
			$worksheetnames[] = $sheet->getTitle();
		}
	}
	if (is_null($foundit)) {
		var_dump($worksheetnames);
	}

	$sheetObj = $objPHPExcel->getSheetByName($foundit);
//	$sheetObj = $objPHPExcel->getActiveSheet();
//	$sheetData = $sheetObj->toArray(null,true,true,true);

	//get rid of empty cells
	$maxCell = $sheetObj->getHighestRowAndColumn();
	$sheetData = $sheetObj->rangeToArray('A1:' . $maxCell['column'] . $maxCell['row'],null,true,true,false);

	$salaryrow = findRow("Salaries and Benefits");
	$materialsrow = findRow("Materials and Supplies");
	$equipmentrow = findRow("Equipment");
	$salesrow = findRow("Cost of Sales");
	$interdivchargesrow = findRow("Interdivisional Charges");
	$otherrow = findRow("Other Expenditures");
	$servicesrow = findRow("Services & Rents");
	$capitalrow = findRow("Contributions to Capital");
	$ccrfrow = findRow("Contributions to CCRF Naming Rights");
	$toreserverow = findRow("Contributions to Reserve/Res Funds");
	if (is_null($toreserverow)) { $toreserverow = findRow("Contributions to Reserves"); }
	$toreservefundrow = findRow("Contributions to Reserve Funds");

	$interdivrecorow = findRow("Interdivisional Recoveries");
	$provsubrow = findRow("Provincial Subsidies");
	$fedsubrow = findRow("Federal Subsidies");
	$othersubrow = findRow("Other Subsidies");
	$userfeesrow = findRow("User Fees & Donations");
	$transcaprow = findRow("Transfers from Capital Fund");
	$fromreserverow = findRow("Contribution from Reserve/Res Funds");
	if (is_null($fromreserverow)) { $fromreserverow = findRow("Contribution from Reserve"); }
	$fromreservefundrow = findRow("Contributions from Reserve Funds");
	if (is_null($fromreservefundrow)) { $fromreservefundrow = findRow("Contribution from Reserve Funds"); }
	$sundryrevrow = findRow("Sundry Revenues");

	$approvedpositionsrow = findRow("APPROVED POSITIONS");

	$actual2009col = findCol("2009","Actual");
	$actual2010col = findCol("2010","Actual");
	$budget2011col = $actual2010col + 1;
	$projected2011col = findCol("2011","Projected");
	$budget2012col = findCol("2012","Budget");
	$outlook2013col = findCol("2013","Outlook");
	$outlook2014col = $outlook2013col + 1;

	$columnvar["2009 Actual"] = "actual2009col";
	$columnvar["2010 Actual"] = "actual2010col";
	$columnvar["2011 Budget"] = "budget2011col";
	$columnvar["2011 Projected"] = "projected2011col";
	$columnvar["2012 Budget"] = "budget2012col";
	$columnvar["2013 Outlook"] = "outlook2013col";
	$columnvar["2014 Outlook"] = "outlook2014col";

//	print "the amount budgeted for salaries in 2012 is $".number_format($sheetData[$salaryrow][$budget2012col]*1000).PHP_EOL;
//	print "approved positions in 2012 ".$sheetData[$approvedpositionsrow][$budget2012col].PHP_EOL;
//	print "average salary in 2012 is $".number_format($sheetData[$salaryrow][$budget2012col]*1000/$sheetData[$approvedpositionsrow][$budget2012col]).PHP_EOL;

	$division = str_replace('.xlsx','',$thisfile);
	$division = str_replace('.xls','',$division);

	if ($outputaslist) {

		//FORMAT:	division , year , actual or projected or budget or outlook , revenue OR expense , type , amount

		foreach ($columnloop as $c) {
			$year = substr($c,0,4);
			$kind = substr($c,5);

			//output revenues
			$revarr = array_keys($revenues);
			for ($i = 0; $i < count($revarr); $i++) {		
				if (isset(${$revarr[$i]})) {
					if (isset($sheetData[${$revarr[$i]}][${$columnvar[$c]}])) {
							file_put_contents($outputrevenuesfilename,sprintf('"%s",%s,%s,%s,%s,%s'.PHP_EOL,$division,$year,$kind,"revenues",$revenues[$revarr[$i]], round($sheetData[${$revarr[$i]}][${$columnvar[$c]}]*1000,2) ),FILE_APPEND | LOCK_EX);
					} else {
	//print 'ZZ ('.$division.')('.$revarr[$i].')('.$c.') isnt set: '.$sheetData[${$revarr[$i]}][${$columnvar[$c]}].PHP_EOL;
					}
				}
			}

			//output expenses		
			$exparr = array_keys($expenses);
			for ($i = 0; $i < count($exparr); $i++) {		
				if (isset(${$exparr[$i]})) {
					if (isset($sheetData[${$exparr[$i]}][${$columnvar[$c]}])) {
							file_put_contents($outputexpensesfilename,sprintf('"%s",%s,%s,%s,%s,%s'.PHP_EOL,$division,$year,$kind,"expenses",$expenses[$exparr[$i]], round($sheetData[${$exparr[$i]}][${$columnvar[$c]}]*1000,2) ),FILE_APPEND | LOCK_EX);
					} else {
	//print '** ('.$division.')('.$exparr[$i].')('.$c.') isnt set: '.$sheetData[${$exparr[$i]}][${$columnvar[$c]}].PHP_EOL;
					}
				}
			}

			//output positions
			file_put_contents($outputpositionsfilename,sprintf('"%s",%s,%s,%s,%s'.PHP_EOL,$division,$year,$kind,"positions", $sheetData[$approvedpositionsrow][${$columnvar[$c]}]) ,FILE_APPEND | LOCK_EX);

		}
	} 
	if ($outputaschart) {

		//always 2012-budget
		//FORMAT: division,salaries,materials,equipment,sales,interdiv,services,other,capital,ccrf,reserve,resfunds
		//file_put_contents($outputchartfilename,"division,salaries,materials,equipment,sales,interdiv,services,other,capital,ccrf,reserve,resfunds".PHP_EOL,FILE_APPEND | LOCK_EX);
		file_put_contents($outputchartfilename,sprintf('"%s",%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s'.PHP_EOL,$division,
			round($sheetData[$salaryrow][$budget2012col]*1000,2),
			round($sheetData[$materialsrow][$budget2012col]*1000,2),
			round($sheetData[$equipmentrow][$budget2012col]*1000,2),
			round($sheetData[$salesrow][$budget2012col]*1000,2),
			round($sheetData[$interdivchargesrow][$budget2012col]*1000,2),
			round($sheetData[$servicesrow][$budget2012col]*1000,2),
			round($sheetData[$otherrow][$budget2012col]*1000,2),
			round($sheetData[$capitalrow][$budget2012col]*1000,2),
			round($sheetData[$ccrfrow][$budget2012col]*1000,2),
			round($sheetData[$toreserverow][$budget2012col]*1000,2),
			round($sheetData[$toreservefundrow][$budget2012col]*1000,2)),
			FILE_APPEND | LOCK_EX);
	}

	if ($outputasjson) {

		unset($divisionarray);

		$c = "2012 Budget";

		$year = substr($c,0,4);
		$kind = substr($c,5);

		//loop through expense types
		$exparr = array_keys($expenses);
		$divisioncount = 0;
		for ($i = 0; $i < count($exparr); $i++) {
			if (isset(${$exparr[$i]})) {
				if ( isset($sheetData[${$exparr[$i]}][${$columnvar[$c]}]) && ($sheetData[${$exparr[$i]}][${$columnvar[$c]}] > 0) )  {
					$divisionarray[$divisioncount]["name"] = $expenses[$exparr[$i]];
					$divisionarray[$divisioncount]["size"] = round($sheetData[${$exparr[$i]}][${$columnvar[$c]}]*1000,2);
//					$divisionarray[$divisioncount]["size"] = "10";
					$divisioncount++;
				} else {
//print '** ('.$division.')('.$exparr[$i].')('.$c.') isnt set: '.$sheetData[${$exparr[$i]}][${$columnvar[$c]}].PHP_EOL;
				}
			} else {
print "something else (".${$exparr[$i]}.")".PHP_EOL;
			}
		}
		//var_dump($divisionarray);

		$budgetarray[$jsoncount]["name"] = $division;
		$budgetarray[$jsoncount]["children"] = $divisionarray;
		$jsoncount++;
print $jsoncount.PHP_EOL;
	}


	unset($sheetData);
	unset($sheetObj);
	unset($objReader);
	unset($objPHPExcel);
}

if ($outputasjson) {
	$object = new stdClass();
	$object->name = "2012 budget";
	$object->children = $budgetarray;

	//var_dump($object);
	//print json_encode($object);
	file_put_contents($outputjsonfilename,json_encode($object));
}

function findRow($search) {
	global $sheetData;
	foreach($sheetData as $rownum => $row) {
		foreach($row as $colnum => $thiscell) {
			if ( strcasecmp($thiscell,$search) == 0 ) {
				return $rownum; // only look for the first occurrence!
			}
		}
	}
}

function findCol($search,$searchbelow) {
	global $sheetData;

	$answer = null;
	foreach($sheetData as $rownum => $row) {
		foreach($row as $colnum => $thiscell) {
			if ( strpos($thiscell,$search) !== false ) {
				if ( strcasecmp($thiscell,$search." ".$searchbelow) == 0 ) {
					$answer = $colnum;
				} elseif ( strcasecmp($thiscell,$search) == 0 ) {
					if ( strpos($sheetData[$rownum+1][$colnum],$searchbelow) !== false ) {
						$answer = $colnum;
					}
				}
			}
		}
	}
	return $answer;
}

?>
