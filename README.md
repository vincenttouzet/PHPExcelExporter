# PHPExcelSourceIterator

Provide a SourceIterator for Excel files for [**sonata-project/exporter**][1]

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
