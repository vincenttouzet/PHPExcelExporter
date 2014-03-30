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
    private $headers = array();

    public function __construct($file, $sheetName = null)
    {
        $objReader = \PHPExcel_IOFactory::createReaderForFile($file);
        //$objReader->setReadDataOnly(true);
        $objPHPExcel = $objReader->load($file);
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
                if ($cell->isFormula()) {
                    $value = $cell->getCalculatedValue();
                }
                $data[$header] = $value;
            }
        } else {
            // get headers
            foreach ($cellIterator as $cell) {
                $value = $cell->getFormattedValue();
                if ($this->hasHeaders && $this->position!==1) {
                    $data[$this->headers[$cell->getColumn()]] = $value;
                } else {
                    $data[$cell->getColumn()] = $value;
                }
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
