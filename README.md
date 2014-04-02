# PHPExcelSourceIterator

[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/vincenttouzet/PHPExcelExporter/badges/quality-score.png?s=e7ae02fbf7a8e3c2e342be8eed2a11316afc36d9)](https://scrutinizer-ci.com/g/vincenttouzet/PHPExcelExporter/)
[![Code Coverage](https://scrutinizer-ci.com/g/vincenttouzet/PHPExcelExporter/badges/coverage.png?s=8cb5e4b2d241d7148d8d3c1dd6ddcd29d536ca87)](https://scrutinizer-ci.com/g/vincenttouzet/PHPExcelExporter/)
[![Build Status](https://travis-ci.org/vincenttouzet/PHPExcelExporter.svg?branch=master)](https://travis-ci.org/vincenttouzet/PHPExcelExporter)

Provide a SourceIterator and a Writer for Excel files for [**sonata-project/exporter**][1]

## Installation

You can use this source with composer:

```json
{
    "require": {
        "sonata-project/exporter": "dev-master",
        "phpoffice/phpexcel": "dev-master",
        "vincet/phpexcel-exporter": "dev-master"
    },
}
```

This library does not require a specific version of [**sonata-project/exporter**][1] or [**phpoffice/phpexcel**][2] so you can use whatever version you want.

## Usage

```php

$source = new VinceT\PHPExcelExporter\Source\PHPExcelSourceIterator(__DIR__.'/file.xlsx');

foreach ($source as $data) {
    ...
}

```

[1]: https://github.com/sonata-project/exporter
[2]: https://github.com/PHPOffice/PHPExcel
