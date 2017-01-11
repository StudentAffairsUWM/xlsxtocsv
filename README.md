xlsxtocsv
=========

An XLSX to CSV parser

## Example Use

```php
$converter = new \XlsxToCsv\XlsxToCsv($filename);
$converter->sheetNomber = 1;
$tmpPath = $converter->convert();
```
