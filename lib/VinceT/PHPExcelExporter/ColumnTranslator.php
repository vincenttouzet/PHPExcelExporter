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

class ColumnTranslator
{
    protected static $chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

    /**
     * Retrieve column index from name
     * e.g: A will return 1, Z will return 26, AB will return 28 ...
     *
     * @param $name
     *
     * @return int
     */
    public static function getColumnIndex($name)
    {
        $chars = str_split($name, 1);
        $index = 0;
        $i = 0;
        while (count($chars) > 0) {
            $c = array_pop($chars);
            $c_index = strpos(self::$chars, $c) + 1;
            if ($i > 0) {
                $index += $c_index * strlen(self::$chars);
            } else {
                $index += $c_index;
            }
            $i++;
        }

        return $index;
    }

    /**
     * Retrieve column name from index
     * e.g: 1 will return A, 26 will return Z, 28 will return AB, ...
     *
     * @param $index
     *
     * @return string
     */
    public static function getColumnName($index)
    {
        $chars = self::$chars;
        $nb = strlen($chars);
        $q = floor($index / $nb);
        $r = $index % $nb;
        $name = '';
        if ($q && $r !== 0) {
            $name .= self::getColumnName($q);
        }
        if ($r) {
            $name .= $chars[$r-1];
        } else {
            $name .= substr($chars, -1, 1);
        }

        return $name;
    }
}
