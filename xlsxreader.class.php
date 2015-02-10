<?php

if (!class_exists("XMLReader")) { Throw New Exception("XMLReader not found"); }
if (!class_exists("ZipArchive")) { Throw New Exception("ZipArchive not found"); }

class XLSXReader
{
	protected $xlsxFile   = "";

	protected $decodeUTF8 = false;

	protected $sheets        = array();
	protected $activeSheet   = 0;
	protected $sharedStrings = array();

	protected $currentXML    = null;

	public function __construct($xlsxFile)
	{
		if (file_exists($xlsxFile) === false)
		{
			Throw New Exception("XLSXReader: XLSX file not found");
		}

		$this->xlsxFile = realpath($xlsxFile);
	}

	public function __destruct()
	{
		if (is_null($this->currentXML) === false)
		{
			$this->currentXML->close();
		}
	}

	public function read()
	{
		$this->readWorkbook();
		$this->readSharedStrings();
	}

	protected function decodeString($string)
	{
		return ($this->decodeUTF8 === true ? utf8_decode($string) : $string);
	}

	public function decodeUTF8($value)
	{
		$this->decodeUTF8 = (is_bool($value) ? $value : true);
	}

	public function getActiveSheet()
	{
		$i = 0;

		foreach ($this->sheets as $sheet)
		{
			if (is_null($this->activeSheet))
			{
				return $sheet["id"];
			}
			elseif ($this->activeSheet == $i)
			{
				return $sheet["id"];
			}

			$i++;
		}
	}

	protected function readWorkbook()
	{
		$xml = new XMLReader();

		if (($result = $xml->open("zip://" . $this->xlsxFile . "#xl/workbook.xml")) === false)
		{
			$xml->close();
			return false;
		}

		while ($xml->read())
		{
			if ($xml->name === "workbookView" && $xml->nodeType == XMLReader::ELEMENT)
			{
				$this->activeSheet = $xml->getAttribute("activeTab");
			}

			if ($xml->name === "sheet" && $xml->nodeType == XMLReader::ELEMENT)
			{
				$rID      = $xml->getAttribute("r:id");
				$isHidden = ($xml->getAttribute("state") == "hidden" ? true : false);

				$this->sheets[$rID] = array (
					"name"   => $this->decodeString($xml->getAttribute("name")),
					"id"     => $rID,
					"hidden" => $isHidden,
					"file"   => ""
				);
			}
		}

		$xml->close();

		$this->readWorkbookRels();
	}

	protected function readWorkbookRels()
	{
		$xml = new XMLReader();

		if (($result = $xml->open("zip://" . $this->xlsxFile . "#xl/_rels/workbook.xml.rels")) === false)
		{
			$xml->close();
			return false;
		}

		while ($xml->read())
		{
			if ($xml->name === "Relationship" && $xml->nodeType == XMLReader::ELEMENT)
			{
				$type = $xml->getAttribute("Type");

				if ($type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
				{
					$id     = $xml->getAttribute("Id");
					$target = $xml->getAttribute("Target");

					if (isset($this->sheets[$id]))
					{
						$this->sheets[$id]["file"] = $target;
					}					
				}
			}
		}

		$xml->close();
	}

	protected function readSharedStrings()
	{
		$xml = new XMLReader();

		if (($result = $xml->open("zip://" . $this->xlsxFile . "#xl/sharedStrings.xml")) === false)
		{
			$xml->close();
			return false;
		}

		while ($xml->read())
		{
			if ($xml->name === "si" && $xml->nodeType == XMLReader::ELEMENT)
			{
				$node = $xml->expand();

				if ($node->hasChildNodes())
				{
					if ($node->childNodes->length == 1 && $node->firstChild->nodeName == "t")
					{
						$this->sharedStrings[] = $node->firstChild->textContent;
					}
					else
					{
						$this->sharedStrings[] = "#";
					}
				}
				else
				{
					$this->sharedStrings[] = "#";
				}
			}
		}

		$xml->close();
	}

	public function getSheets()
	{
		return $this->sheets;
	}

	public function getSheetRow($sheetID)
	{
		if (isset($this->sheets[$sheetID]) === false)
			return false;

		$sheet = $this->sheets[$sheetID];

		if (is_null($this->currentXML))
		{
			$this->currentXML = new XMLReader();

			$xml = &$this->currentXML;

			if (($result = $xml->open("zip://" . $this->xlsxFile . "#xl/" . $sheet["file"])) === false)
			{
				$xml->close();
				return false;
			}

			# Move to first row
			while ($xml->read() && $xml->name !== "row");
		}
		else
		{
			$xml = &$this->currentXML;
			
			if ($xml->next("row") === false)
			{
				$xml->close();
				$this->currentXML = null;

				return array();
			}
		}

		$cols = array();

		if ($xml->name === "row")
		{
			/**
			* attribut 't'
			* - b for boolean
			* - d for date
			* - inlineStr for an inline string
			* - n for number
			* - s for shared string
			* - str for a formula
			*/
			$node = $xml->expand();

			if ($node->hasChildNodes() === true)
			{
				foreach ($node->childNodes as $cellNode)
				{
					# For each cell 'c'
					if ($cellNode->nodeName === "c")
					{
						$cellType  = $cellNode->getAttribute("t");
						$vTags     = $cellNode->getElementsByTagName("v");

						$cellNum   = $cellNode->getAttribute("r");
						$cellNum   = preg_replace("/[0-9]+$/", "", $cellNum);
						$cellNum   = $this->excelLettersToNum($cellNum);
						$cellNum   = $cellNum - 1;

						$cellValue = "";

						foreach ($vTags as $vTag)
						{
							$cellValue = $vTag->nodeValue;
						}

						if (empty($cellType))
						{
							if (($num = filter_var($cellValue, FILTER_VALIDATE_FLOAT)) !== false)
							{
								$cellValue = $num;
							}
						}
						elseif ($cellType === "s") # sharedStrings
						{
							$cellValue = (isset($this->sharedStrings[$cellValue]) ? $this->sharedStrings[$cellValue] : "#");
						}

						$cols[$cellNum] = $this->decodeString($cellValue);
					}
				}
			}
		}
		else
		{
			$xml->close();
			$this->currentXML = null;
		}

		return $cols;
	}

	public function getSheetDatas($sheetID)
	{
		if (isset($this->sheets[$sheetID]) === false)
			return false;

		$sheet = $this->sheets[$sheetID];
		
		$xml = new XMLReader();

		if (($result = $xml->open("zip://" . $this->xlsxFile . "#xl/" . $sheet["file"])) === false)
		{
			$xml->close();
			return false;
		}

		# Move to first row
		while ($xml->read() && $xml->name !== "row");

		$rows = array();

		while ($xml->name === "row")
		{
			$cols = array();

			/**
			* attribut 't'
			* - b for boolean
			* - d for date
			* - inlineStr for an inline string
			* - n for number
			* - s for shared string
			* - str for a formula
			*/
			$node = $xml->expand();

			if ($node->hasChildNodes() === true)
			{
				foreach ($node->childNodes as $cellNode)
				{
					# For each cell 'c'
					if ($cellNode->nodeName === "c")
					{
						$cellType  = $cellNode->getAttribute("t");
						$vTags     = $cellNode->getElementsByTagName("v");

						$cellNum   = $cellNode->getAttribute("r");
						$cellNum   = preg_replace("/[0-9]+$/", "", $cellNum);
						$cellNum   = $this->excelLettersToNum($cellNum);
						$cellNum   = $cellNum - 1;

						$cellValue = "";

						foreach ($vTags as $vTag)
						{
							$cellValue = $vTag->nodeValue;
						}

						if (empty($cellType))
						{
							if (($num = filter_var($cellValue, FILTER_VALIDATE_FLOAT)) !== false)
							{
								$cellValue = $num;
							}
						}
						elseif ($cellType === "s") # sharedStrings
						{
							$cellValue = (isset($this->sharedStrings[$cellValue]) ? $this->sharedStrings[$cellValue] : "#");
						}

						$cols[$cellNum] = $this->decodeString($cellValue);
					}
				}
			}

			$rows[] = $cols;		

			$xml->next("row");
		}

		$xml->close();

		return $rows;
	}

	protected function excelLettersToNum($letters)
	{
		$num = 0;
		$arr = array_reverse(str_split($letters));

		for ($i = 0; $i < count($arr); $i++)
		{
			$num += (ord(strtolower($arr[$i])) - 96) * (pow(26, $i));
		}

		return $num;
	}
}

?>