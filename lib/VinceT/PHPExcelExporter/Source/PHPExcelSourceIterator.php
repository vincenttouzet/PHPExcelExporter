<?php

/*
 * This file is part of the php-excel-exporter package.
 *
 * (c) Vincent Touzet <vincent.touzet@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace VinceT\PHPExcelExporter\Source;

use Exporter\Source\SourceIteratorInterface;

class PHPExcelSourceIterator implements SourceIteratorInterface
{
    /**
     * @var \PHPExcel_Worksheet_RowIterator
     */
    private $rowIterator = null;
    private $sheet = null;
    private $position = 1;
    private $hasHeaders = true;
    private $getIndexedCells = null;
    private $headers = array();

    /**
     * @param string      $file
     * @param bool        $hasHeaders      Must be true if the first line contains column name
     * @param null|string $sheetName       If null use active sheet
     * @param bool        $getIndexedCells only used when hasHeaders is false.
     *                                     If getIndexedCells is true data will be an associative array
     */
    public function __construct($file, $hasHeaders = true, $sheetName = null, $getIndexedCells = true)
    {
        $objReader = \PHPExcel_IOFactory::createReaderForFile($file);
        //$objReader->setReadDataOnly(true);
        $objPHPExcel = $objReader->load($file);
        $this->hasHeaders = $hasHeaders;
        $this->getIndexedCells = $getIndexedCells;
        $this->sheet = $objPHPExcel->getActiveSheet();
        if ($sheetName) {
            $this->sheet = $objPHPExcel->getSheetByName($sheetName);
        }
        $this->rowIterator = $this->sheet->getRowIterator();
    }

    /**
     * {@inheritdoc}
     */
    public function current()
    {
        $row = $this->rowIterator->current();
        $cellIterator = $row->getCellIterator();
        $data = array();
        if ($this->hasHeaders && $this->position !== 1) {
            $data = array_flip($this->headers);
            foreach ($data as &$val) {
                $val = '';
            }
        }
        /* @var \PHPExcel_Cell $cell */
        if ($this->hasHeaders && $this->position !== 1) {
            // get data
            foreach ($this->headers as $column => $header) {
                $cell = $this->sheet->getCell($column.$this->position);
                $value = $cell->getFormattedValue();
                $data[$header] = $value;
            }
        } else {
            // get headers
            $index = 1;
            foreach ($cellIterator as $cell) {
                $value = $cell->getFormattedValue();
                if ($this->getIndexedCells) {
                    $data[$cell->getColumn()] = $value;
                } else {
                    $data[$index] = $value;
                }
                $index++;
            }
        }

        return $data;
    }

    /**
     * {@inheritdoc}
     */
    public function next()
    {
        $this->rowIterator->next();
        $this->position++;
    }

    /**
     * {@inheritdoc}
     */
    public function key()
    {
        return $this->position;
    }

    /**
     * {@inheritdoc}
     */
    public function valid()
    {
        return $this->rowIterator->valid();
    }

    /**
     * {@inheritdoc}
     */
    public function rewind()
    {
        $this->rowIterator->rewind();
        $this->position = 1;
        if ($this->hasHeaders && $this->valid()) {
            $this->headers = $this->current();
            $this->next();
        }
    }
}
