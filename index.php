<?php

error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');
set_time_limit(0);

define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');

$pic_name='logo.png';
$im = ImageCreateFromPng($pic_name);
$localinfo = getimagesize($pic_name);
$height=$localinfo[1];
$width=$localinfo[0];
function get_rgb($im,$x,$y){
	
	$rgb = ImageColorAt($im, $x, $y);
	$r = ($rgb >> 16) & 0xFF;
	$g = ($rgb >> 8) & 0xFF;
	$b = $rgb & 0xFF;
	return new_dechex($r).new_dechex($g).new_dechex($b);
	//return $rgb;

}

function new_dechex($num){
	$num2=dechex($num);
	if($num<17)
		$num2='0'.$num2;
	return $num2;
}



/** Include PHPExcel */
require_once './Classes/PHPExcel.php';



$objPHPExcel = new PHPExcel();

$objPHPExcel->setActiveSheetIndex(0);

$objActSheet = $objPHPExcel->getActiveSheet();
$objActSheet->setTitle('测试Sheet');

for($w=0;$w<$width;$w++){
	for($h=0;$h<$height;$h++){
		$objStyleA5 = $objActSheet->getStyle(getCellName($w).$h);
		$objFillA5 = $objStyleA5->getFill();
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue(getCellName($w).$h, get_rgb($im,$w,$h));
		$objFillA5->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
		$objFillA5->getStartColor()->setRGB(get_rgb($im,$w,$h));
	
	}

}



$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save(str_replace('.php', '.xlsx', __FILE__));



function getCellName($num){
	$str='';
	$arr=array('0'=>'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

	do{
	 $d=$num%26;
	 $num=floor($num/26);
	 $str=$arr[$d].$str;
	}while($num>0);

	return $str;
}


