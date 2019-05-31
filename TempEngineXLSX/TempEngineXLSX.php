<?php

class TempEngineXLSX {
	private $contentArr;  // содержит содержимое и пути файлов xml
	private $zip;         // для организации нового архива
	private $newFileName; // имя нового файла
	private $newFileDir;  // путь к новому файлу
	private $tmpDir;      // директория работы
	private $sheets;      // пути к файлам страницы и изх содержимое
	private $dict;		  // массив содержащий словарь (значения стринг в ячейках)
	private $rowsArr;     // массив, содержащай структуру строк, сформированных с из файла sheets 
	private $data;




	/* Конструктор класса
	 * string $src - Путь к шаблону
	 * string $dst - Путь к новову xlsx файлу
	 * array $data - Массив данных
	 */
	public function __construct($src, $dst, $data) {
		$this->zip  = new ZipArchive();
		$this->data = $data;
		$this->rowsArr = [];
		$this->newFileName = basename($dst, "xlsx");
		$dirDst = dirname($dst);
		if(!file_exists($dirDst)) {
			mkdir($dirDst);
		}
		$this->newFileDir = $dirDst;
		$newFile =$this->newFileDir."\\tmp.zip";
		copy($src, $newFile);
		$this->tmpDir = $dirDst."\\"."archive";
		$zipTemplate = new ZipArchive();
		$zipTemplate->open($newFile);
		$zipTemplate->extractTo($this->tmpDir);
		$zipTemplate->close();
		$sheetsFilesArr = glob($this->tmpDir."\\xl\\worksheets\\*.xml");
		foreach ($sheetsFilesArr as $value) {
			$sheetsArr[] = 
				[
					"path"  => $value,
					"value" => file_get_contents($value),
				];
		}
		
		$this->contentArr["styles"]            = 
			[
				"path"  => $this->tmpDir."\\xl\\styles.xml",
				"value" => file_get_contents($this->tmpDir."\\xl\\styles.xml"),
			];
		$this->contentArr["workbook"]          = 
			[
				"path"  => $this->tmpDir."\\xl\\workbook.xml",
				"value" => file_get_contents($this->tmpDir."\\xl\\workbook.xml")
			];
		$this->contentArr["sharedStrings.xml"] = 
			[
				"path"  => $this->tmpDir."\\xl\\sharedStrings.xml",
				"value" => file_get_contents($this->tmpDir."\\xl\\sharedStrings.xml")
			];
		$this->contentArr["workbook.xml.rels"] = 
			[
				"path"  => $this->tmpDir."\\xl\\_rels\\workbook.xml.rels",
				"value" => file_get_contents($this->tmpDir."\\xl\\_rels\\workbook.xml.rels")
			];
		$this->contentArr[".rels"]             = 
			[
				"path"  => $this->tmpDir."\\_rels\\.rels",
				"value" => file_get_contents($this->tmpDir."\\_rels\\.rels")
			];
		$this->sheets = $sheetsArr;
	}
	
	/* Функция удаляющая папку
	 * 
	 */
	private function clean_folder($__dir){
		$filesArr = glob($__dir."\\*");
		foreach ($filesArr as $file) {
			if(is_dir($file)){
				$this->clean_folder($file);
			}
			else {
				unlink($file);
			}
		}
		rmdir($__dir);
	}
	
	/*Обобщающая функция создания нового xlsx
	 * 
	 */

	public function getNewXLSX() {
		
		$this->readDict();
		$count = count($this->sheets);
		for($index = 0; $index < $count; ++$index) {
			$this->readRows($index);
			$this->simpleFill();
			
			
			while ($founded = $this->searchArray()) {
				if($founded === "g") {
					$this->fillGorizontal();
				}
				else {
					$this->fillVertical();
				}
			}
			
			
			$this->formWorkSheet($index);
			foreach($this->rowsArr as $key => $row) {
				//var_dump($row->code);
			}
		}
		
		$this->formSharedStrings();
		$this->putContent();
		$this->createXLSX();
		
		unlink($this->tmpDir."\\_rels\\.rels");
		$this->clean_folder($this->tmpDir);
		unlink($this->newFileDir."\\tmp.zip");

	}

	/*
	 * Функция считывающая строки с файла sheets
	 */
	private function readRows($nomer_sheet) {
		preg_match_all("#<row[^<>]+/>#", $this->sheets[$nomer_sheet]["value"], $match);
		preg_match_all("#<row[^/]+?>.*?</row>#", $this->sheets[$nomer_sheet]["value"], $match1);
		$needArr = array_merge($match[0], $match1[0]);
		foreach ($needArr as $value) {
			$obj = new Row($value);
			$this->rowsArr[$obj->index] = $obj;
		}
		ksort($this->rowsArr);
	}
	
	
	/*
	 * Функциия формирует оновленный файл sheets
	 */
	private function formWorkSheet($nomer_sheet) {
		$newRowsCode = "";
		foreach ($this->rowsArr as $row) {
			$newRowsCode .= $row->code;
		}
		
		$this->sheets[$nomer_sheet]["value"] = 
			preg_replace("#<row.*</row>#", $newRowsCode, $this->sheets[$nomer_sheet]["value"] );
	}
	
	/*
	 * Функция считывающая словарь
	 */
	private function readDict() {
		preg_match_all("#<t>(.+?)</t>#",$this->contentArr["sharedStrings.xml"]["value"], $match);
		$this->dict = $match[1];
	}
	
	/*
	 * Функция, добавляющее нговое значение в словарь
	 */
	private function addToDict($__newString) {
		if(!array_search($__newString, $this->dict)){
			$this->dict[] = $__newString;
		}
		return array_search($__newString, $this->dict);
	}
	
	/*
	 * Функция формирует файл с словарем 
	 */
	private function formSharedStrings() {
		$str = "";
		foreach ($this->dict as $value) {
			$str .= "<si><t>$value</t></si>";
		}
		$this->contentArr["sharedStrings.xml"]["value"] = 
			preg_replace("#<si>.*</si>#", $str, $this->contentArr["sharedStrings.xml"]["value"]);
	}
	
	
	/*
	 * Функция, осуществляющая поиск ячеек со значением, которое соответствует
	 *  либо горизонтальному массиву либо вертикальному.  
	 */
	private function searchArray() {
		foreach ($this->rowsArr as $indexRow => $row) {
			foreach($row->cellsArr as $indexCell => $cell) {
				if(!$cell->is_string) continue;
				$value = $this->dict[$cell->value];
				if(preg_match("#^\w+?_v$#", $value,$m) && isset($this->data[$m[0]]) ) {
					return "v";
				}
				elseif(preg_match("#^\w+?_g$#", $value,$m) && isset($this->data[$m[0]]) ){
					return "g";
				}
			}
		}
		
		return false;
	}

	/* 
	 * Функция вертикальной вставки
	 */
	private function fillVertical(){
		$AllCountNewRows = 0;
		foreach($this->rowsArr as $indexRow => $row) {
			$countNewRows = 0;
			foreach($row->cellsArr as $indexCell => $cell) {
				
				if($cell->is_string === false) continue;
				$realValueOfCell =  $this->dict[$cell->value];
				if(!preg_match("#^\w+?_v$#", $realValueOfCell, $match)) continue;
				$variableName = $match[0];
				if(!isset($this->data[$variableName])){
					continue;
				}
				$count = count($this->data[$variableName]);
				
				$index = 0;
				while($index < $count) {
					
					$element = $this->data[$variableName][$index];
					$value = is_string($element) 
						? $this->addToDict($element)
						: $element;
					
					preg_match("#(\w+?)(\d+)#", $indexCell, $arr);
					$newKeyCell = $arr[1].($arr[2] + $index);
					if($index == 0 || $index <= $countNewRows) {
						
						 $this->rowsArr[$indexRow + $index + $AllCountNewRows]
							 ->editCell($newKeyCell, is_string($element), $value);
					}
					else {
						$this->copyRow(0, $indexRow + $index + $AllCountNewRows - 1 , true);
						$countNewRows++;
						$this->rowsArr[$indexRow + $index + $AllCountNewRows]
							 ->editCell($newKeyCell, is_string($element), $value);
					}
				++$index;
				}	
			}
			$AllCountNewRows += $countNewRows;
		}
	}
	/*
	 * Функция горизонтальной вставки
	 */
	private function fillGorizontal(){
		foreach($this->rowsArr as $indexRow => $row) {
			$countNewCells = 0;
			foreach($row->cellsArr as $indexCell => $cell) {
				
				if($cell->is_string === false) continue;
				$realValueOfCell =  $this->dict[$cell->value];
				if(!preg_match("#^\w+?_g$$#", $realValueOfCell, $match)) continue;
				$variableName = $match[0];
				if(!isset($this->data[$variableName])){
					continue;
				}
				$count = count($this->data[$variableName]);
				
				$index = 0;
				preg_match("#(\w)(\d+)#", $indexCell, $m);
				$letter = ord($m[1]) + $countNewCells;
				while($index < $count) {
					$element = $this->data[$variableName][$index];
					$value = is_string($element) 
						? $this->addToDict($element)
						: $element;
					
					$newID = chr($letter + $index).$m[2];
					$prevID = chr($letter + $index - 1).$m[2];
					if($index == 0) {
						 $this->rowsArr[$indexRow]->editCell($newID, is_string($element), $value);
					}
					else {
						$this->rowsArr[$indexRow]->copyCell($prevID);
						$this->rowsArr[$indexRow]->editCell($newID, is_string($element), $value);
						++$countNewCells;
					}
					++$index;	
				}
			}
		}
	}
	
	/*
	 * Поиск и замена линейных значений
	 */
	private function simpleFill() {
		foreach ($this->data as $key => $value){
			if(!is_array($value)){
				$this->simpleReplace($key);
			}
		}
	}
	
	/*
	 * Замена линейных значений
	 */
	private function simpleReplace($__key) {
		$replacement = "{".$__key."}";
		
		if(is_string($this->data[$__key])) {
			foreach($this->dict as $index => $string){
				$this->dict[$index] = preg_replace("#$replacement#", $this->data[$__key], $string);
			}
		}
		else {
			foreach($this->dict as $index => $string){
				if(!preg_match("#^$replacement$#", $this->dict[$index])) {
					$this->dict[$index] = preg_replace("#$replacement#", $this->data[$__key], $string);
				}
			}
			
			$index = array_search($replacement, $this->dict);
			if($index === false) return;
			foreach ($this->rowsArr as $key1 => $rowObj){
				foreach ($this->rowsArr[$key1]->cellsArr as $key2 => $cellObj){
					
					if($cellObj->value === $index && $cellObj->is_string) {
						$rowObj->editCell($key2, false, $this->data[$__key]);
					}
				}
			}	
		}	
	}
	
	
	/*
	 * Копирование строки
	 */
	private function copyRow($nomer_sheet, $nomer_row, $empty = false) {
		
		$newRow = new Row ( $this->rowsArr[$nomer_row]->code);
		$newRow->setIndex($newRow->index + 1);
		
		$tmpArr = [];
		foreach($this->rowsArr as $key => $rowObj) {
			if($key < $nomer_row + 1)				continue;
			$rowObj->setIndex($rowObj->index + 1);
			unset($this->rowsArr[$key]);
			$tmpArr[$key + 1] = $rowObj;
		}
		$this->rowsArr = array_replace($this->rowsArr,$tmpArr);
		
		
		if($empty) {
			foreach($newRow->cellsArr as $keyCell => $cellObj) {
				$newRow->editCell($keyCell, false, false);
			}
		}
		
		$this->rowsArr[$newRow->index] = $newRow;
		ksort($this->rowsArr);
		return 0;
	}
	
	/*
	 * обновить все файлы, содержащиеся в content
	 */
	private function putContent(){
		foreach($this->contentArr as $key => $data) {
			file_put_contents($data["path"], $data["value"]);
		}
		foreach($this->sheets as $key => $data) {
			file_put_contents($data["path"], $data["value"]);
		}
	}
	
	/*
	 * создание xlsx-документа
	 */
	private function createXLSX() {
		$fullPath = $this->newFileDir."\\".$this->newFileName.".zip";
		$this->zip->open($fullPath, ZipArchive::CREATE);
		$this->createZip($this->tmpDir);
		$this->zip->addFile($this->tmpDir."\\_rels\\.rels", "_rels\\.rels" );
		$this->zip->close();
		rename($fullPath, $this->newFileDir."\\".$this->newFileName."xlsx");
	}
	
	/*
	 * Организация zip-архива, для создания xlsx
	 */
	private function createZip($pathDir){
		$filesArr = glob($pathDir."\\*");
		foreach($filesArr as $file) {
			$nameFile = basename($file);
			if(is_dir($file)) {
				$this->createZip($file);
			}
			else {
				$dir = str_replace($this->tmpDir."\\", "", dirname($file)."\\");
				$dir = $dir === $this->tmpDir."\\" ? "" : $dir; 
				$this->zip->addFile($file, $dir.$nameFile);
			}
		}
	}
	
	public function download($file) {
		if (file_exists($file)) {
		  // сбрасываем буфер вывода PHP, чтобы избежать переполнения памяти выделенной под скрипт
		  // если этого не сделать файл будет читаться в память полностью!
		  if (ob_get_level()) {
			ob_end_clean();
		  }
		  // заставляем браузер показать окно сохранения файла
		  header('Content-Description: File Transfer');
		  header('Content-Type: application/octet-stream');
		  header('Content-Disposition: attachment; filename=' . basename($file));
		  header('Content-Transfer-Encoding: binary');
		  header('Expires: 0');
		  header('Cache-Control: must-revalidate');
		  header('Pragma: public');
		  header('Content-Length: ' . filesize($file));
		  // читаем файл и отправляем его пользователю
		  readfile($file);
		  exit;
		}
	}

}
/*
 * Вспомогательный класс для хранения строки
 */
class Row {
	public $index;
	public $code;
	public $cellsArr;
	public	function __construct($__string) {
		$this->code = $__string;
		preg_match('#r="([\d]+)"#', $__string, $match2);
		$this->index = (int)$match2[1];
		
		preg_match_all("#<c.+?/(?:>|c>)#", $__string, $match1);
		$this->cellsArr = [];
		foreach($match1[0] as $code) {
			$obj = new Cell($code);
			$this->cellsArr[$obj->id] = $obj;
		}
	}
	
	public function copyCell($__id){
		$newCell = new Cell($this->cellsArr[$__id]->code);
		preg_match('#(\w)(\d+)#', $newCell->id, $match);
		$newId = chr(ord($match[1]) + 1).$match[2];
		$newCell->setID($newId);
		$newCell->editCell(false, false); 
		$tmpArr = [];
		foreach($this->cellsArr as $key => $cellObj) {
			if($key <= $__id) continue;
			preg_match('#(\w)(\d+)#', $cellObj->id, $match);
			$curID = chr(ord($match[1]) + 1).$match[2];
			$cellObj->setId($curID);
			unset($this->cellsArr[$key]);
			$tmpArr[$curID] = $cellObj;
		}
		$this->cellsArr = array_replace($this->cellsArr, $tmpArr);
		$this->cellsArr[$newId] = $newCell;
		ksort($this->cellsArr);
		$this->ungradeCode();
	}
	
	private function ungradeCode(){
		$newCellsCode = "";
		foreach ($this->cellsArr as $cell) {
			$newCellsCode .= $cell->code;
		}
		preg_match('#<row.*?>#', $this->code, $match);
		$tail = preg_match('#</row>#', $this->code)? "</row>" : "";
		$this->code = $match[0].$newCellsCode.$tail;
	}
	
	public function setIndex($newValue) {
		$this->index = $newValue;
		$this->code = 
			preg_replace('#r="\d+"#', 'r="'.$newValue.'"',$this->code);
		foreach ($this->cellsArr as $key =>$cell) {
			$cell->setID($newValue);
			unset($this->cellsArr[$key]);
			$this->cellsArr[$cell->id] = $cell;
		}
		
		$this->ungradeCode();
		return 0;
	}
	
	public function editCell($__cellIndex, $__is_string, $__value) {
		if(!isset($this->cellsArr[$__cellIndex])){
			var_dump("НЕТ Такой  $__cellIndex");
		}
		
		if($__value !== false) {
			$this->cellsArr[$__cellIndex]->is_string = $__is_string;
		}
		$this->cellsArr[$__cellIndex]->value = $__value;
		$this->cellsArr[$__cellIndex]->code = '<c r="'.$this->cellsArr[$__cellIndex]->id.'"';
		
		$style = $this->cellsArr[$__cellIndex]->style === false
			?""
			:' s="'.$this->cellsArr[$__cellIndex]->style.'"';
		$is_string = $this->cellsArr[$__cellIndex]->is_string === false
			?""
			:' t="s"';
		
		$this->cellsArr[$__cellIndex]->code .= $__value === false 
			? "$style/>"
			: "$style$is_string><v>$__value</v></c>";
		
		$this->ungradeCode();
	}
}
/*
 * Вспомогательный класс для хранения ячейки
 */
class Cell {
	public $id;
	public $code;
	public $value;
	public $is_string;
	public $style;
	public function __construct($__string) {
		//можно и в одну регулярку.
		$this->code = $__string;
		preg_match('#r="([\w\d]+)"#', $__string, $match1);
		$this->id = $match1[1];
		$this->is_string = preg_match('#t="s"#', $__string) ? true : false;
		preg_match('#<v>(\d+)</v>#', $__string, $match2);
		$this->value = isset($match2[1]) ? (int)$match2[1] : false;
		preg_match('#s="([\d]+)"#', $__string, $match3);
		$this->style = isset($match3[1]) ? $match3[1] : false;
	}
	
	public function setID($newValue){
		if(is_string($newValue)) {
			$this->id = preg_replace("#[\w\d]+#", $newValue, $this->id);
		}
		else {
			$this->id =  preg_replace("#\d+#", $newValue, $this->id);
		}
		$this->code = preg_replace('#r="[\w\d]+"#','r="'.$this->id.'"',$this->code);
	}
	
	public function editCell($__is_string, $__value){
		if($__value !== false) {
			$this->is_string = $__is_string;
		}
		$this->value = $__value;
		$this->code = '<c r="'.$this->id.'"';
		
		$style = $this->style === false
			?""
			:' s="'.$this->style.'"';
		$is_string = $this->is_string === false
			?""
			:' t="s"';
		
		$this->code .= $__value === false 
			? "$style/>"
			: "$style$is_string><v>$__value</v></c>";
	}
}