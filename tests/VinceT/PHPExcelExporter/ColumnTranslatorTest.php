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

class ColumnTranslatorTest extends \PHPUnit_Framework_TestCase
{
    protected $map = array(
        'A' => 1,
        'E' => 5,
        'Z' => 26,
        'AA' => 27,
        'BA' => 53,
    );

    public function testGetColumnName()
    {
        foreach ($this->map as $name => $index) {
            $this->assertEquals($name, ColumnTranslator::getColumnName($index));
        }
    }

    public function testGetColumnIndex()
    {
        foreach ($this->map as $name => $index) {
            $this->assertEquals($index, ColumnTranslator::getColumnIndex($name));
        }
    }
}
