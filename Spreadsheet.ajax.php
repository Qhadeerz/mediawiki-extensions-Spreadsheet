<?php

class SpreadsheetAjax {
	public static function getData($file,$sheetIndex){
		$file = wfLocalFile($file);
		$sheetIndex = is_numeric($sheetIndex) ? intval($sheetIndex) : 0;
		if($file->exists()){

			$readerType = self::getReaderType($file->getExtension());
			$reader = PHPExcel_IOFactory::createReader($readerType);

			if($reader instanceof PHPExcel_Reader_Excel2007){
				$sheetNames = $reader->listWorksheetNames($file->getLocalRefPath());
				$sheetName = $sheetNames[$sheetIndex];
				$reader->setLoadSheetsOnly($sheetName);
				$reader->setIncludeCharts(true);
				$reader->setReadDataOnly(false);
			}
			$phpexcel = $reader->load($file->getLocalRefPath());
			$phpexcel->setActiveSheetIndex(0);

			$writer = new PHPExcel_Writer_JSON($phpexcel);
			$output = $writer->toJson();
			$phpexcel->disconnectWorksheets();
			return $output;
		}
		http_response_code(404);	
	}

	private static function getReaderType($fileExtension){
		switch (strtolower($fileExtension)) {
			case 'csv':
				return 'CSV';
			case 'xls':
				return 'Excel5';
			case 'xlsx':
				return 'Excel2007';
			case 'ods':
				return 'OOCalc';
			default:
				return '';
		}
	}
}
