<?php

/**
 * Author: Manuel Martinez
 * email:tiempor3al.manuel@gmail.com 
 * date: December 2017
 * Licence: MIT
 */

namespace tiempor3al\Excel;


class Reader{

    private $filename;
    private $unique_id;
    private $sharedStrings = null;
    private $relationships = [];
    private $worksheets = [];


    const END_ROW = 'end_row';
    const START_ROW = 'start_row';


    /**
     * @param $filename
     * @throws \Exception
     */
    function Initialize($filename){

        if(file_exists($filename) == false){
            throw new \Exception(sprintf("%s file could not be found",$filename));
        }

        $this->filename = $filename;
        $this->unique_id = uniqid();

        $this->ExtractWorkbookRelationships();
        $this->ExtractWorksheets();
        $this->ExtractSharedStrings();

    }

    /**
     *
     */
    private function ExtractSharedStrings()
    {

        $file = sprintf("zip://%s#xl/sharedStrings.xml",$this->filename);

        $reader = new \XMLReader();
        $reader->open($file);

        $i = 0;
        $value = null;

        $sharedStrings = new \SplFixedArray(0);

        while($reader->read()){

            if ($reader->nodeType == \XMLReader::ELEMENT && $reader->name == 'si'){

                $sharedStrings->setSize($i + 1);
                $sharedStrings[$i]  = $reader->readString();
                $i++;

            }
        }
        $reader->close();
        $this->sharedStrings = $sharedStrings;
    }


    /**
     *
     */
    private function ExtractWorkbookRelationships()
    {

        $file = sprintf("zip://%s#xl/_rels/workbook.xml.rels",$this->filename);

        $reader = new \XMLReader();
        $reader->open($file);


        while($reader->read()){

            if ($reader->nodeType == \XMLReader::ELEMENT && $reader->name == 'Relationship'){

                $this->relationships[$reader->getAttribute('Id')] = $reader->getAttribute('Target');
            }
        }
        $reader->close();

    }

    /**
     *
     */
    private function ExtractWorksheets()
    {

        $file = sprintf("zip://%s#xl/workbook.xml",$this->filename);

        $reader = new \XMLReader();
        $reader->open($file);


        while($reader->read()){

            if ($reader->nodeType == \XMLReader::ELEMENT && $reader->name == 'sheet'){

                $this->relationships[$reader->getAttribute('Id')] = $reader->getAttribute('Target');

                $id = $reader->getAttributeNs('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                $name = trim($reader->getAttribute('name'));

                $this->worksheets[$name] = $id;
            }
        }
        $reader->close();

    }


    /**
     * @param $worksheetName
     * @param array $cells
     * @return \Generator
     * @throws \Exception
     */
    public function ParseWorksheet($worksheetName, $cells = [])
    {
        if(!isset($this->worksheets[$worksheetName])){
            throw new \Exception(sprintf("%s worksheet could not be found",$worksheetName));
        }

        $id = $this->worksheets[$worksheetName];

        $worksheet = $this->relationships[$id];

        $file = sprintf("zip://%s#xl/%s",$this->filename,$worksheet);

        $reader = new \XMLReader();
        $reader->open($file);

        $currentCell = null;

        $hasRanges = false;
        foreach($cells as $range){
            if(strpos($range,':') !== false){
                $hasRanges = true;
                break;
            }
        }


        while($reader->read()){

            if ($reader->nodeType == \XMLReader::ELEMENT && $reader->name == 'row'){
                yield self::START_ROW;
            }

            if ($reader->nodeType == \XMLReader::END_ELEMENT && $reader->name == 'row'){
                yield self::END_ROW;
            }

            if ($reader->nodeType == \XMLReader::ELEMENT && $reader->name == 'c'){

                $currentCell = $reader->getAttribute('r');
                $currentCol = preg_replace('/[0-9]+/', '', $currentCell);
                $currentRow = preg_replace('/[A-Z]+/', '', $currentCell);

                $cellType = $reader->getAttribute('t');

                $value = null;
                //Shared strings
                if($cellType !== null && $cellType == 's'){
                    $sharedIndex = $reader->readString();
                    $value = $this->sharedStrings[$sharedIndex];
                }


                //Dates
                if($cellType !== null && $cellType == 'n'){
                    $value = $reader->readString();
                }

                if($cellType !== null && $cellType == 'str'){

                }

                if($cellType === null) {
                    $value = $reader->readString();
                }

                if(in_array($currentCol,$cells)) {
                    yield array('cell' => $currentCell, 'row' => $currentRow, 'col' => $currentCol, 'value' => $value);
                }

                if($hasRanges) {
                    foreach ($cells as $range) {
                        if (strpos($range, ':') !== false) {
                            $range_values = explode(':', $range);
                            $low = $this->getColumnNumber($range_values[0]);
                            $high = $this->getColumnNumber($range_values[1]);

                            $currentColNumber = $this->getColumnNumber($currentCol);

                            if ($currentColNumber >= $low && $currentColNumber <= $high) {
                                yield array('cell' => $currentCell, 'row' => $currentRow, 'col' => $currentCol, 'value' => $value);
                            }
                        }
                    }
                }

            }

        }
        $reader->close();

    }


    /**
     * @param $column
     * @return int
     */
    public function getColumnNumber($column){

        $length = strlen($column);

        $sum = 0;
        for($i = 0; $i < $length; $i++){

            $sum *= 26;
            $sum += (ord($column[$i]) - ord('A') + 1);
        }

        return $sum;
    }

    /**
     * @return mixed
     */
    public function getSharedStringsSize()
    {
        return $this->sharedStrings->getSize();
    }

    /**
     * @return array
     */
    public function getWorksheets(){
        return array_keys($this->worksheets);
    }

    /**
     * @return DateTime
     */
    public function toDate($value)
    {
        //Internally dates are stored as days since January 1st, 1900
        $date = new \DateTime('1900-01-01');
        //For historic reasons Excel adds two days 
        //see http://www.kirix.com/stratablog/excel-date-conversion-days-from-1900
        //So we need to substract them
        $days = intval($value) - 2;
        $interval = sprintf("P%sD",$days);
        return $date->add(new \DateInterval($interval));
    }


}
