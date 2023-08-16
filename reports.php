<?php
include 'odbfunc.php';
require_once __DIR__ . '\PHPExcel.php';
require_once __DIR__ . '\PHPExcel\Writer\Excel2007.php';
 
class Reports
{
		public static function rep6ATTU() // Отчет о поездках по маршруту 6А ТТУ за вчерашний день
		{
			$td = date("Y-m-d H:i:s", mktime(0, 0, 0, date("m"), date("d"), date("Y")));
			$ytd = date("Y-m-d H:i:s", mktime(0, 0, 0, date("m"), date("d")-1, date("Y")));
			$tmrw = date("Y-m-d H:i:s", mktime(0, 0, 0, date("m"), date("d")+1, date("Y")));
			$shytd = date("Y-m-d", mktime(0, 0, 0, date("m"), date("d")-1, date("Y")));
			$res = odbFunc::trip6ATTU($td, $ytd, $tmrw);
			if ($res != 0)
			{
				$xls = new PHPExcel();
				$xls->getProperties()->setTitle("Отчет о поездках");
				$xls->getProperties()->setCompany("Трамвайно-троллейбусное управление");
				$xls->setActiveSheetIndex(0);
				$sheet = $xls->getActiveSheet();
				$sheet->setTitle('Лист 1');
				$sheet->getColumnDimension("A")->setWidth(17,6);
				$sheet->getColumnDimension("B")->setWidth(17,6);
				$sheet->getColumnDimension("C")->setWidth(17,6);
				$sheet->getColumnDimension("D")->setWidth(17,6);
				$sheet->getColumnDimension("E")->setWidth(17,6);
				$sheet->getColumnDimension("F")->setWidth(17,6);
				$sheet->getColumnDimension("G")->setWidth(17,6);
				$sheet->getColumnDimension("H")->setWidth(17,6);
			
				$titlestyle = array
				(
					'font' => array
					(
						'name'      => 'Microsoft Sans Serif',
						'size'      => 12,   
						'bold'      => true,
					)
				);
				
				$tabheadstyle = array
				(
					'font' => array
					(
						'name'      => 'Microsoft Sans Serif',
						'size'      => 10,   
						'bold'      => true,
					)
				);
				
				$maintextstyle = array
				(
					'font' => array
					(
						'name'      => 'Microsoft Sans Serif',
						'size'      => 10,   
						'bold'      => false,
					)
				);
				
				$border = array
				(
					'borders'=>array
					(
						'allborders' => array
						(
							'style' => PHPExcel_Style_Border::BORDER_THIN,
							'color' => array('rgb' => '000000')
						)
					)
				);

 				$sheet->getStyle("A6:H7")->applyFromArray($border);
				
				$sheet->getStyle('C1:C2')->applyFromArray($titlestyle);
				$sheet->setCellValue("C1", "Отчет о поездках ");
				$sheet->setCellValue("C2", "Трамвайно-троллейбусное управление");

				$sheet->getStyle('A4')->applyFromArray($tabheadstyle);
				$sheet->setCellValue("A4", "Дата:");
		
				$sheet->getStyle('B4')->applyFromArray($maintextstyle);
				$sheet->getStyle("B4")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);
				$sheet->setCellValue("B4", $shytd);
				
				$sheet->getStyle('A6:H6')->applyFromArray($tabheadstyle);
				$sheet->getStyle("A6:H6")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
				$sheet->setCellValue("A6", "Наличные");
				$sheet->setCellValue("B6", "Банковские карты");
				$sheet->setCellValue("C6", "ЕТК");
				$sheet->setCellValue("D6", "Безлимитные ЕТК");
				$sheet->setCellValue("E6", "Карты школьника");
				$sheet->setCellValue("F6", "Карты студента");
				$sheet->setCellValue("G6", "Социальные карты");
				$sheet->setCellValue("H6", "ИТОГО");
				
				$sheet->getStyle('A7:H7')->applyFromArray($maintextstyle);
				$sheet->getStyle("A7:H7")->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

				$sheet->setCellValue("A7", $res[0]["snal"]);
				$sheet->setCellValue("B7", $res[0]["sbank"]);
				$sheet->setCellValue("C7", $res[0]["setk"]);
				$sheet->setCellValue("D7", $res[0]["sbetk"]);
				$sheet->setCellValue("E7", $res[0]["sksh"]);
				$sheet->setCellValue("F7", $res[0]["sks"]);
				$sheet->setCellValue("G7", $res[0]["ssoc"]);
				$sum = $res[0]["snal"]+$res[0]["sbank"]+$res[0]["setk"]+$res[0]["sbetk"]+$res[0]["sksh"]+$res[0]["sks"]+$res[0]["ssoc"];
				$sheet->setCellValue("H7", $sum);

				$objWriter = new PHPExcel_Writer_Excel2007($xls);
				$objWriter->save(__DIR__ .'\temp\ТТУ_маршрут_6А_'.$shytd.'.xlsx');
				$repname = __DIR__ .'\temp\ТТУ_маршрут_6А_'.$shytd.'.xlsx';
				return $repname;
			}
			else
			{
				return 0;
			}
		}
	
	
		




}

?>