# MNB PHPExcel XML

Independent streaming XML reader and writer. Requires core and `ext-libxml`.

```bash
composer require mnb/mnb-phpexcel-xml:^2.0
```

```php
use Mnb\PHPExcel\Format\Xml;

$rows = Xml::read('customers.xml')->withHeaderRow()->toArray();
Xml::write($rows, 'customers-export.xml');
```

`ext-xmlreader` is recommended for native forward-only streaming; the core compatibility parser remains available when it is missing.
