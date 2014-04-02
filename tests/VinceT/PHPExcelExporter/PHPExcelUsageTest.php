<?php

/*
 * This file is part of the php-excel-exporter package.
 *
 * (c) Vincent Touzet <vincent.touzet@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace VinceT\PHPExcelExporter;

use VinceT\PHPExcelExporter\Source\PHPExcelSourceIterator;
use VinceT\PHPExcelExporter\Writer\PHPExcelWriter;

class PHPExcelUsageTest extends \PHPUnit_Framework_TestCase
{
    protected $data = array(
        0 => array(
            'username' => 'johndoe',
            'firstname' => 'John 1',
            'name' => 'Doe',
        ),
        1 => array(
            'username' => 'johndoe_2',
            'firstname' => 'John 2',
            'name' => 'Doe',
        ),
    );

    public function testUsage()
    {
        $filename = tempnam(sys_get_temp_dir(), uniqid()).'usage.xlsx';
        $writer = new PHPExcelWriter($filename);
        $this->write($writer);
        $source = new PHPExcelSourceIterator($filename);
        $this->doTest($source, true, true);
        unlink($filename);
    }

    public function testSheetName()
    {
        $filename = tempnam(sys_get_temp_dir(), uniqid()).'sheetName.xlsx';
        $writer = new PHPExcelWriter($filename, true, 'Excel2007', 'My custom sheet name');
        $this->write($writer);
        $source = new PHPExcelSourceIterator($filename, true, 'My custom sheet name');
        $this->doTest($source, true, true);
        unlink($filename);
    }

    public function testNoHeaders()
    {
        $filename = tempnam(sys_get_temp_dir(), uniqid()).'noheaders.xlsx';
        $writer = new PHPExcelWriter($filename, false, 'Excel2007', 'My custom sheet name');
        $this->write($writer);
        $source = new PHPExcelSourceIterator($filename, false, 'My custom sheet name');
        $this->doTest($source, false, true);
        unlink($filename);
    }

    public function testNoHeadersNoIndexedCells()
    {
        $filename = tempnam(sys_get_temp_dir(), uniqid()).'noheaders.xlsx';
        $writer = new PHPExcelWriter($filename, false, 'Excel2007', 'My custom sheet name');
        $this->write($writer);
        $source = new PHPExcelSourceIterator($filename, false, 'My custom sheet name', false);
        $this->doTest($source, false, false);
        unlink($filename);
    }

    protected function write(PHPExcelWriter $writer)
    {
        $data = $this->data;
        $writer->open();
        $writer->write($data[0]);
        $writer->write($data[1]);
        $writer->close();
    }

    public function doTest(PHPExcelSourceIterator $source, $hasHeaders, $indexedCells)
    {
        $data_source = array();
        $begin = 1;
        if ($hasHeaders) {
            $begin++;
        }
        foreach ($source as $k => $d) {
            $this->assertEquals($begin, $k);
            $data_source[] = $d;
            $begin++;
        }

        foreach ($this->data as $key => $values) {
            $index = 1;
            foreach ($values as $k => $v) {
                if ($hasHeaders) {
                    $this->assertEquals($v, $data_source[$key][$k]);
                } else {
                    if ($indexedCells) {
                        $this->assertEquals($v, $data_source[$key][ColumnTranslator::getColumnName($index)]);
                    } else {
                        $this->assertEquals($v, $data_source[$key][$index]);
                    }
                }
                $index++;
            }
        }
    }
}
