<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Format;

use Mnb\PHPExcel\Reader\Options\ReaderOptions;
use Mnb\PHPExcel\Reader\ReadSession;
use Mnb\PHPExcel\Reader\XmlReader;
use Mnb\PHPExcel\Writer\XmlWriter;

final class Xml
{
    /** @param array<string,mixed>|ReaderOptions $options */
    public static function read(string $path, array|ReaderOptions $options = []): ReadSession
    {
        return new ReadSession($path, new XmlReader(), $options);
    }

    /** @param iterable<array<int|string,mixed>|mixed> $rows @param array<string,mixed> $options */
    public static function write(iterable $rows, string $path, array $options = []): void
    {
        $buffer = [];
        foreach ($rows as $row) {
            $buffer[] = is_array($row) ? $row : [$row];
        }
        (new XmlWriter())->writeRows($buffer, $path, $options);
    }
}
