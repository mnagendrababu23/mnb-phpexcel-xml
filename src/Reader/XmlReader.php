<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\Arr;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\LocaleNormalizer;
use Mnb\PHPExcel\Support\MnbExcelException;
use Mnb\PHPExcel\Support\Xml\XmlReader as NativeXmlReader;

/** Secure, forward-only XML row reader with optional schema mapping. */
final class XmlReader implements IterableReaderInterface
{
    /** @return list<list<mixed>> */
    public function readSheet(string $path, int|string $sheet = 1, array $options = []): array
    {
        return array_values(iterator_to_array($this->iterateSheet($path, $sheet, $options), true));
    }

    /** @return \Generator<int,list<mixed>> */
    public function iterateSheet(string $path, int|string $sheet = 1, array $options = []): iterable
    {
        $projection = ColumnProjection::fromOptions($options);
        $rows = (function () use ($path, $sheet, $options, $projection): \Generator {
            $columns = isset($options['xml_columns']) ? array_values(array_map('strval', (array) $options['xml_columns'])) : [];
            $includeHeader = (bool) ($options['xml_header_row'] ?? $options['include_header_row'] ?? true);
            $headerYielded = false;
            $output = 0;
            $extraKeys = strtolower((string) ($options['streaming_extra_keys'] ?? 'ignore'));
            if (!in_array($extraKeys, ['ignore', 'error'], true)) {
                throw new MnbExcelException('streaming_extra_keys must be "ignore" or "error".');
            }

            foreach ($this->iterateRawRows($path, $sheet, $options) as $row) {
                $row = $projection->project($row);
                if (Arr::isAssoc($row)) {
                    if ($columns === []) {
                        $columns = array_map('strval', array_keys($row));
                    } elseif ($extraKeys === 'error') {
                        $unknown = array_values(array_diff(array_map('strval', array_keys($row)), $columns));
                        if ($unknown !== []) {
                            throw MnbExcelException::withCode(
                                'XML row contains columns not present in the streaming schema: ' . implode(', ', $unknown),
                                ErrorCode::FILE_READ_FAILED,
                                ['columns' => $unknown]
                            );
                        }
                    }
                    if ($includeHeader && !$headerYielded) {
                        yield $output++ => $columns;
                        $headerYielded = true;
                    }
                    $line = [];
                    foreach ($columns as $column) {
                        $line[] = $row[$column] ?? null;
                    }
                    yield $output++ => $line;
                    continue;
                }

                if ($includeHeader && $columns !== [] && !$headerYielded) {
                    yield $output++ => $columns;
                    $headerYielded = true;
                }
                yield $output++ => array_values($row);
            }
        })();

        yield from $this->sliceIterable($rows, $options);
    }

    /** @return array<string,list<array<int|string,mixed>>> */
    public function readWorkbook(string $path, array $options = []): array
    {
        $workbook = [];
        foreach ($this->iterateRawRowsWithSheet($path, null, $options) as $item) {
            $workbook[$item['sheet']][] = $item['row'];
        }
        if ($workbook === []) {
            $defaultSheetName = trim((string) ($options['sheet_name'] ?? 'Sheet1')) ?: 'Sheet1';
            $workbook[$defaultSheetName] = [];
        }
        return $workbook;
    }

    /** @return list<string> */
    public function sheetNames(string $path, array $options = []): array
    {
        $this->ensureExtension();
        $this->assertReadableFile($path, $options);
        $schema = XmlSchemaMapping::from($options['xml_schema'] ?? $options['schema'] ?? null, $options);
        $sheetTag = $schema?->sheetTag() ?? (string) ($options['sheet_tag'] ?? 'sheet');
        $nameAttribute = $schema?->sheetNameAttribute() ?? 'name';
        $defaultSheetName = trim((string) ($options['sheet_name'] ?? 'Sheet1')) ?: 'Sheet1';
        $reader = $this->openReader($path);
        $names = [];
        $sawSheet = false;
        try {
            while ($reader->read()) {
                if ($reader->nodeType === NativeXmlReader::DOC_TYPE && !(bool) ($options['allow_doctype'] ?? false)) {
                    throw MnbExcelException::withCode('XML document types are disabled for secure reading.', ErrorCode::SECURITY_BLOCKED, ['path' => $path]);
                }
                if ($reader->nodeType === NativeXmlReader::ELEMENT && $reader->localName === $sheetTag) {
                    $sawSheet = true;
                    $base = trim((string) ($reader->getAttribute($nameAttribute) ?? ''));
                    $names[] = $this->uniqueSheetNameFromList($names, $base !== '' ? $base : 'Sheet' . (count($names) + 1));
                }
            }
        } finally {
            $reader->close();
        }
        return $sawSheet ? $names : [$defaultSheetName];
    }

    /** @return \Generator<int,array<int|string,mixed>> */
    private function iterateRawRows(string $path, int|string $sheet, array $options): \Generator
    {
        foreach ($this->iterateRawRowsWithSheet($path, $sheet, $options) as $index => $item) {
            yield $index => $item['row'];
        }
    }

    /**
     * @return \Generator<int,array{sheet:string,row:array<int|string,mixed>,source_row:int}>
     */
    private function iterateRawRowsWithSheet(string $path, int|string|null $sheet, array $options): \Generator
    {
        $this->ensureExtension();
        $this->assertReadableFile($path, $options);
        $schema = XmlSchemaMapping::from($options['xml_schema'] ?? $options['schema'] ?? null, $options);
        $sheetTag = $schema?->sheetTag() ?? (string) ($options['sheet_tag'] ?? 'sheet');
        $rowTag = $schema?->rowTag() ?? (string) ($options['row_tag'] ?? $options['row'] ?? 'row');
        $nameAttribute = $schema?->sheetNameAttribute() ?? 'name';
        $defaultSheetName = trim((string) ($options['sheet_name'] ?? 'Sheet1')) ?: 'Sheet1';
        $allowDoctype = (bool) ($options['allow_doctype'] ?? false);
        $maxSourceRows = isset($options['max_source_rows']) ? max(0, (int) $options['max_source_rows']) : null;
        $selectedIndex = $sheet === null ? null : (is_int($sheet) || ctype_digit((string) $sheet) ? max(1, (int) $sheet) : null);
        $selectedName = $sheet !== null && is_string($sheet) && !ctype_digit($sheet) ? $sheet : null;

        $previousLibxmlMode = libxml_use_internal_errors(true);
        libxml_clear_errors();
        $reader = $this->openReader($path);
        $currentSheet = $defaultSheetName;
        $sheetDepth = null;
        $sheetCount = 0;
        $sawSheetTag = false;
        $totalRows = 0;
        $selectedRows = 0;
        $knownNames = [];
        $matchedSheet = $sheet === null;
        $xmlErrors = [];

        try {
            while ($reader->read()) {
                if ($reader->nodeType === NativeXmlReader::DOC_TYPE && !$allowDoctype) {
                    throw MnbExcelException::withCode('XML document types are disabled for secure reading.', ErrorCode::SECURITY_BLOCKED, ['path' => $path]);
                }

                if ($reader->nodeType === NativeXmlReader::ELEMENT && $reader->localName === $sheetTag) {
                    $sawSheetTag = true;
                    $sheetCount++;
                    $base = trim((string) ($reader->getAttribute($nameAttribute) ?? ''));
                    $currentSheet = $this->uniqueSheetNameFromList($knownNames, $base !== '' ? $base : 'Sheet' . $sheetCount);
                    $knownNames[] = $currentSheet;
                    $sheetDepth = $reader->depth;
                    continue;
                }
                if ($reader->nodeType === NativeXmlReader::END_ELEMENT && $reader->localName === $sheetTag && $sheetDepth === $reader->depth) {
                    $sheetDepth = null;
                    $currentSheet = $defaultSheetName;
                    continue;
                }
                if ($reader->nodeType !== NativeXmlReader::ELEMENT || $reader->localName !== $rowTag) {
                    continue;
                }
                if (!$sawSheetTag) {
                    $currentSheet = $defaultSheetName;
                    $sheetCount = 1;
                } elseif ($sheetDepth === null) {
                    continue;
                }

                $totalRows++;
                if ($maxSourceRows !== null && $totalRows > $maxSourceRows) {
                    throw MnbExcelException::withCode(
                        'XML row limit exceeded. Rows: ' . $totalRows . ', max_source_rows: ' . $maxSourceRows,
                        ErrorCode::FILE_READ_FAILED,
                        ['path' => $path, 'rows' => $totalRows, 'max_source_rows' => $maxSourceRows]
                    );
                }

                $matches = $sheet === null
                    || ($selectedIndex !== null && $sheetCount === $selectedIndex)
                    || ($selectedName !== null && $currentSheet === $selectedName);
                if (!$matches) {
                    $this->skipCurrentElement($reader);
                    continue;
                }
                $matchedSheet = true;

                $rowXml = $reader->readOuterXml();
                if ($rowXml === '') {
                    continue;
                }
                $parsed = $this->parseRowXml($rowXml, $options);
                $row = $schema?->map($parsed['row'], $parsed['attributes'], $parsed['text']) ?? $parsed['row'];
                $selectedRows++;
                yield $selectedRows - 1 => ['sheet' => $currentSheet, 'row' => $row, 'source_row' => $totalRows];
            }
        } finally {
            $reader->close();
            $xmlErrors = libxml_get_errors();
            libxml_clear_errors();
            libxml_use_internal_errors($previousLibxmlMode);
        }

        foreach ($xmlErrors as $xmlError) {
            if ($xmlError->level >= LIBXML_ERR_ERROR) {
                throw MnbExcelException::withCode(
                    'Invalid XML near line ' . $xmlError->line . ': ' . trim($xmlError->message),
                    ErrorCode::FILE_READ_FAILED,
                    ['path' => $path, 'line' => $xmlError->line, 'column' => $xmlError->column]
                );
            }
        }
        if (!$matchedSheet && $sheet !== null) {
            throw new MnbExcelException('XML sheet does not exist: ' . (string) $sheet);
        }
    }

    private function skipCurrentElement(NativeXmlReader $reader): void
    {
        if ($reader->isEmptyElement) {
            return;
        }
        $depth = $reader->depth;
        while ($reader->read()) {
            if ($reader->nodeType === NativeXmlReader::END_ELEMENT && $reader->depth === $depth) {
                return;
            }
        }
    }

    /** @return array{row:array<int|string,mixed>,attributes:array<string,string>,text:string} */
    private function parseRowXml(string $xml, array $options): array
    {
        $reader = new NativeXmlReader();
        if (!@$reader->XML($xml, null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode('Unable to parse XML row.', ErrorCode::FILE_READ_FAILED);
        }
        try {
            while ($reader->read() && $reader->nodeType !== NativeXmlReader::ELEMENT) {
            }
            if ($reader->nodeType !== NativeXmlReader::ELEMENT) {
                return ['row' => [], 'attributes' => [], 'text' => ''];
            }
            $attributes = [];
            if ($reader->hasAttributes && $reader->moveToFirstAttribute()) {
                do {
                    $attributes[$reader->localName] = $reader->value;
                } while ($reader->moveToNextAttribute());
                $reader->moveToElement();
            }

            $rowDepth = $reader->depth;
            $row = [];
            $rowText = '';
            while ($reader->read()) {
                if ($reader->nodeType === NativeXmlReader::END_ELEMENT && $reader->depth === $rowDepth) {
                    break;
                }
                if (($reader->nodeType === NativeXmlReader::TEXT || $reader->nodeType === NativeXmlReader::CDATA) && $reader->depth === $rowDepth + 1) {
                    $rowText .= $reader->value;
                    continue;
                }
                if ($reader->nodeType !== NativeXmlReader::ELEMENT || $reader->depth !== $rowDepth + 1) {
                    continue;
                }
                $name = $reader->localName;
                $column = $reader->getAttribute('column');
                $value = $this->readElementValue($reader, $options, 1);
                if ($name === (string) ($options['cell_tag'] ?? $options['cell'] ?? 'cell') && $column !== null && ctype_digit($column)) {
                    $row[max(0, (int) $column - 1)] = $value;
                    continue;
                }
                if (array_key_exists($name, $row)) {
                    if (!is_array($row[$name]) || Arr::isAssoc($row[$name])) {
                        $row[$name] = [$row[$name]];
                    }
                    $row[$name][] = $value;
                } else {
                    $row[$name] = $value;
                }
            }

            if ($row !== [] && !Arr::isAssoc($row)) {
                ksort($row);
                $max = max(array_keys($row));
                $dense = [];
                for ($i = 0; $i <= $max; $i++) {
                    $dense[] = $row[$i] ?? null;
                }
                $row = $dense;
            }
            return ['row' => $row, 'attributes' => $attributes, 'text' => $rowText];
        } finally {
            $reader->close();
        }
    }

    private function readElementValue(NativeXmlReader $reader, array $options, int $depth): mixed
    {
        $maxDepth = max(1, (int) ($options['max_depth'] ?? 64));
        if ($depth > $maxDepth) {
            throw MnbExcelException::withCode('XML nesting exceeds max_depth: ' . $maxDepth, ErrorCode::FILE_READ_FAILED, ['max_depth' => $maxDepth]);
        }
        if ($reader->isEmptyElement) {
            return null;
        }
        $elementDepth = $reader->depth;
        $children = [];
        $text = '';
        while ($reader->read()) {
            if ($reader->nodeType === NativeXmlReader::END_ELEMENT && $reader->depth === $elementDepth) {
                break;
            }
            if ($reader->nodeType === NativeXmlReader::ELEMENT && $reader->depth === $elementDepth + 1) {
                $name = $reader->localName;
                $value = $this->readElementValue($reader, $options, $depth + 1);
                if (array_key_exists($name, $children)) {
                    if (!is_array($children[$name]) || Arr::isAssoc($children[$name])) {
                        $children[$name] = [$children[$name]];
                    }
                    $children[$name][] = $value;
                } else {
                    $children[$name] = $value;
                }
                continue;
            }
            if (($reader->nodeType === NativeXmlReader::TEXT || $reader->nodeType === NativeXmlReader::CDATA || $reader->nodeType === NativeXmlReader::SIGNIFICANT_WHITESPACE) && $reader->depth === $elementDepth + 1) {
                $text .= $reader->value;
            }
        }
        if ($children !== []) {
            return $children;
        }
        $value = (bool) ($options['trim_values'] ?? false) ? trim($text) : $text;
        if ((bool) ($options['empty_strings_to_null'] ?? false) && $value === '') {
            return null;
        }
        return $this->inferValue($value, $options);
    }

    private function inferValue(string $value, array $options): mixed
    {
        if (!(bool) ($options['infer_types'] ?? false)) {
            return $value;
        }
        $lower = strtolower(trim($value));
        if ($lower === 'true') {
            return true;
        }
        if ($lower === 'false') {
            return false;
        }
        $trimmed = trim($value);
        if (preg_match('/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/', $trimmed) === 1) {
            return LocaleNormalizer::parseCanonicalNumber(
                $trimmed,
                $options,
                false,
                $value,
                (bool) ($options['preserve_leading_zero_numbers'] ?? true)
            );
        }
        return $value;
    }

    /** @param iterable<int,list<mixed>> $rows @return \Generator<int,list<mixed>> */
    private function sliceIterable(iterable $rows, array $options): \Generator
    {
        $startRow = max(1, (int) ($options['start_row'] ?? 1));
        $endRow = isset($options['end_row']) ? max(1, (int) $options['end_row']) : null;
        $sourceSkipRows = max(0, (int) ($options['source_skip_rows'] ?? 0));
        $sourceLimitRows = isset($options['source_limit_rows']) ? max(0, (int) $options['source_limit_rows']) : null;
        $seen = 0;
        $yielded = 0;
        foreach ($rows as $row) {
            $seen++;
            if ($seen < $startRow || $seen <= $sourceSkipRows) {
                continue;
            }
            if ($endRow !== null && $seen > $endRow) {
                break;
            }
            if ($sourceLimitRows !== null && $yielded >= $sourceLimitRows) {
                break;
            }
            yield $seen - 1 => $row;
            $yielded++;
        }
    }

    private function uniqueSheetNameFromList(array $names, string $base): string
    {
        $name = $base;
        $suffix = 2;
        while (in_array($name, $names, true)) {
            $name = $base . '_' . $suffix++;
        }
        return $name;
    }

    private function openReader(string $path): NativeXmlReader
    {
        $reader = new NativeXmlReader();
        if (!@$reader->open($path, null, LIBXML_NONET | LIBXML_COMPACT)) {
            throw MnbExcelException::withCode('Unable to open XML file: ' . $path, ErrorCode::FILE_OPEN_FAILED, ['path' => $path]);
        }
        return $reader;
    }

    private function assertReadableFile(string $path, array $options): void
    {
        if (!is_file($path)) {
            throw MnbExcelException::withCode('XML file not found: ' . $path, ErrorCode::FILE_NOT_FOUND, ['path' => $path]);
        }
        $maxBytes = isset($options['max_file_bytes']) ? max(0, (int) $options['max_file_bytes']) : null;
        $size = filesize($path);
        if ($maxBytes !== null && $size !== false && $size > $maxBytes) {
            throw MnbExcelException::withCode(
                'XML file exceeds max_file_bytes. Size: ' . $size . ', max_file_bytes: ' . $maxBytes,
                ErrorCode::FILE_READ_FAILED,
                ['path' => $path, 'size_bytes' => $size, 'max_file_bytes' => $maxBytes]
            );
        }
    }

    private function ensureExtension(): void
    {
        if (!class_exists(NativeXmlReader::class)) {
            throw MnbExcelException::withCode('ext-xmlreader is required to read XML files.', ErrorCode::EXTENSION_MISSING);
        }
    }
}
