<?php

/*
 * This file is part of the php-excel-exporter package.
 *
 * (c) Vincent Touzet <vincent.touzet@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace VinceT\PHPExcelExporter\Writer;

use Exporter\Writer\WriterInterface;
use VinceT\PHPExcelExporter\ColumnTranslator;

class PHPExcelWriter implements WriterInterface
{
    protected $objPhpExcel = null;
    protected $filename = null;
    protected $showHeaders = null;
    protected $excelType = null;
    protected $sheetName = null;
    protected $position = null;

    /**
     * @param string $filename
     * @param bool   $showHeaders
     * @param string $excelType
     * @param string $sheetName
     *
     * @throws \RuntimeException if filename already exists
     */
    public function __construct($filename, $showHeaders = true, $excelType = 'Excel2007', $sheetName = 'Worksheet')
    {
        $this->objPhpExcel = new \PHPExcel();
        $this->filename = $filename;
        $this->showHeaders = $showHeaders;
        $this->excelType = $excelType;
        $this->sheetName = $sheetName;
        $this->position = 1;

        if (is_file($filename)) {
            throw new \RuntimeException(sprintf('The file %s already exist', $filename));
        }
    }

    /**
     * @return void
     */
    public function open()
    {
        $this->objPhpExcel->getActiveSheet()->setTitle($this->sheetName);
    }

    /**
     * @return void
     */
    public function close()
    {
        $objWriter = \PHPExcel_IOFactory::createWriter($this->objPhpExcel, $this->excelType);
        $objWriter->save($this->filename);
    }

    /**
     * @param array $data
     *
     * @return void
     */
    public function write(array $data)
    {
        if ($this->position == 1 && $this->showHeaders) {
            $this->addHeaders($data);
            $this->position++;
        }
        $this->addLine($data);
        $this->position++;
    }

    /**
     * Get headers to write
     *
     * @param array $data
     */
    protected function addHeaders(array $data)
    {
        $headers = array();
        foreach ($data as $header => $value) {
            $headers[] = $header;
        }
        $this->addLine($headers);
    }

    /**
     * Write a line
     *
     * @param $data
     */
    protected function addLine($data)
    {
        $index = 1;
        foreach ($data as $value) {
            $columnName = ColumnTranslator::getColumnName($index);
            $this->objPhpExcel->getActiveSheet()->getCell($columnName.$this->position)->setValue($value);
            $index++;
        }
    }
}
