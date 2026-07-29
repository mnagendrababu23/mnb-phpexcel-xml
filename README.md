# MNB PHPExcel XML

<<<<<<< HEAD
Streaming XML reader, schema mapping, and XML writer module for MNB PHPExcel.
Documentation URL: https://mnbphpexcel.space/getting-started/installation
This package is generated from the MNB PHPExcel monorepo. Do not copy source files between modules manually.

## Install
=======
Independent streaming XML reader and writer. Requires core and `ext-libxml`.
>>>>>>> f82a541 (Release v2.0.0)

```bash
composer require mnb/mnb-phpexcel-xml:^2.0
```

```php
use Mnb\PHPExcel\Format\Xml;

$rows = Xml::read('customers.xml')->withHeaderRow()->toArray();
Xml::write($rows, 'customers-export.xml');
```

`ext-xmlreader` is recommended for native forward-only streaming; the core compatibility parser remains available when it is missing.
