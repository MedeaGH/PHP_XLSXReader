PHP_XLSXReader
==============

XLSX Excel Spreadsheet Reader in PHP

Simple example:
```php
$reader = new XLSXReader("output.xlsx");
$reader->decodeUTF8(true);
$reader->read();

# Get the sheet list
$sheets = $reader->getSheets();

foreach ($sheets as $sheet)
{
	$datas = $reader->getSheetDatas($sheet["id"]);

	# Do stuff...
}
```
