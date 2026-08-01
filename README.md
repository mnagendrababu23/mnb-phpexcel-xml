# MNB PHPExcel XML

Independent streaming XML reader and writer. Requires core and `ext-libxml`.
## MNB PHPExcel Assistant

Generate MNB PHPExcel code using our dedicated ChatGPT assistant:

[Open MNB PHPExcel AI Assistant](https://chatgpt.com/g/g-6a6e31d80350819194b68853d41c1561-mnb-phpexcel-assistant)
```bash
composer require mnb/mnb-phpexcel-xml:^2.0
```

```php
use Mnb\PHPExcel\Format\Xml;

$rows = Xml::read('customers.xml')->withHeaderRow()->toArray();
Xml::write($rows, 'customers-export.xml');
```

`ext-xmlreader` is recommended for native forward-only streaming; the core compatibility parser remains available when it is missing.
