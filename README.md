Reader is a very simple and fast Excel (.xlsx) parser.

```php

<?php

require __DIR__ . '/vendor/autoload.php';

use tiempor3al\Excel\Reader as ExcelReader;

if(!isset($argv[1])){
    die("Missing Excel filename...");    
}

$reader = new ExcelReader();
$reader->Initialize($argv[1]);
$worksheets = $reader->getWorksheets();

//We parse the first worksheet and the first two columns
$generator = $reader->ParseWorksheet($worksheets[0], ['A:B']);

//Since the library uses generators, it consumes little memory
foreach ($generator as $cell) {
   
    if ($cell === ExcelReader::START_ROW) {
        //Do initialization stuff here
        printf("Start of row...\r\n");
    }

    //if cell is an array then we get its value
    if (is_array($cell)) {

        //$cell holds four values:
        //'cell' => A8 - the position of the cell
        //'row' => 8 - the row of the cell
        //'col' => A - the column of the cell
        //'value' => xyz - the value of the cell

        //For values like strings and numbers, we can get the value in $cell['value']
        if ($cell['col'] == 'A') {
            printf("Cell value: %s\r\n",$cell['value']);
        }
        
        //For dates we need to convert the values using the toDate method
        //Suppose that column B is filled with dates
        if ($cell['col'] == 'B') {
            //toDate returns a DateTime Object
            $date = $reader->toDate($cell['value']);
            
            printf("Cell value: %s\r\n",$date->format('Y-m-d'));
        }
    }

    if ($cell === ExcelReader::END_ROW) {
    	printf("End of row...\r\n");
    }
}


```
