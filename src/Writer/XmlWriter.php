<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Writer;

use Mnb\PHPExcel\Core\CellValue;
use Mnb\PHPExcel\Core\WorkbookData;
use Mnb\PHPExcel\Core\WorksheetData;
use Mnb\PHPExcel\Support\AtomicFileWriter;
use Mnb\PHPExcel\Support\ErrorCode;
use Mnb\PHPExcel\Support\MnbExcelException;

final class XmlWriter
{
    /** @param array<string,mixed> $options */
    public function writeWorkbook(WorkbookData $workbook, string $path, array $options = []): void
    {
        $this->writeString($path, $this->workbookToString($workbook, $options));
    }

    /** @param array<string,mixed> $options */
    public function workbookToString(WorkbookData $workbook, array $options = []): string
    {
        $root = $this->tag((string) ($options['root'] ?? (count($workbook->sheets) === 1 ? 'rows' : 'workbook')));
        $sheetTag = $this->tag((string) ($options['sheet_tag'] ?? 'sheet'));
        $includeDeclaration = (bool) ($options['declaration'] ?? true);
        $includeMetadata = (bool) ($options['include_metadata'] ?? false);
        $indent = (bool) ($options['pretty'] ?? true);

        $xml = $includeDeclaration ? '<?xml version="1.0" encoding="UTF-8"?>' . "\n" : '';

        if (count($workbook->sheets) === 1 && !$includeMetadata && ($options['workbook_wrapper'] ?? false) === false) {
            return $xml . $this->rowsXml($this->sheetRows($workbook->sheets[0], $options), $options + ['root' => $root, 'pretty' => $indent]);
        }

        $xml .= '<' . $root . '>' . "\n";

        if ($includeMetadata && $workbook->metadata !== []) {
            $xml .= $this->valueNode('metadata', $workbook->metadata, 1, $options);
        }

        foreach ($workbook->sheets as $sheet) {
            $xml .= $this->pad(1, $indent) . '<' . $sheetTag . ' name="' . $this->escape($sheet->name) . '">' . "\n";
            $xml .= $this->rowsInnerXml($this->sheetRows($sheet, $options), 2, $options);
            $xml .= $this->pad(1, $indent) . '</' . $sheetTag . '>' . "\n";
        }

        $xml .= '</' . $root . '>' . "\n";
        return $xml;
    }

    /** @param array<string,mixed> $options @return list<array<string,mixed>|list<mixed>> */
    private function sheetRows(WorksheetData $sheet, array $options): array
    {
        return $sheet->rowsForStructuredExport(
            preserveAssociative: (bool) ($options['preserve_associative_rows'] ?? true),
            dataOnly: (bool) ($options['data_only'] ?? false)
        );
    }

    /**
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param array<string,mixed> $options
     */
    public function writeRows(array $rows, string $path, array $options = []): void
    {
        $this->writeString($path, $this->rowsToString($rows, $options));
    }

    /**
     * @param list<array<string,mixed>|list<mixed>> $rows
     * @param array<string,mixed> $options
     */
    public function rowsToString(array $rows, array $options = []): string
    {
        $includeDeclaration = (bool) ($options['declaration'] ?? true);
        $root = $this->tag((string) ($options['root'] ?? 'rows'));
        $xml = $includeDeclaration ? '<?xml version="1.0" encoding="UTF-8"?>' . "\n" : '';
        return $xml . $this->rowsXml($rows, $options + ['root' => $root]);
    }

    /** @param list<array<string,mixed>|list<mixed>> $rows @param array<string,mixed> $options */
    private function rowsXml(array $rows, array $options): string
    {
        $root = $this->tag((string) ($options['root'] ?? 'rows'));
        $indent = (bool) ($options['pretty'] ?? true);
        $xml = '<' . $root . '>' . "\n";
        $xml .= $this->rowsInnerXml($rows, 1, $options);
        $xml .= '</' . $root . '>' . "\n";
        return $xml;
    }

    /** @param list<array<string,mixed>|list<mixed>> $rows @param array<string,mixed> $options */
    private function rowsInnerXml(array $rows, int $level, array $options): string
    {
        $rowTag = $this->tag((string) ($options['row_tag'] ?? $options['row'] ?? 'row'));
        $cellTag = $this->tag((string) ($options['cell_tag'] ?? $options['cell'] ?? 'cell'));
        $indent = (bool) ($options['pretty'] ?? true);
        $includeRowNumber = (bool) ($options['include_row_number'] ?? false);
        $xml = '';

        foreach (array_values($rows) as $rowIndex => $row) {
            if (!is_array($row)) {
                $row = [$row];
            }

            $attrs = $includeRowNumber ? ' number="' . ($rowIndex + 1) . '"' : '';
            $xml .= $this->pad($level, $indent) . '<' . $rowTag . $attrs . '>' . "\n";

            foreach ($row as $key => $value) {
                if (is_int($key)) {
                    $xml .= $this->valueNode($cellTag, $value, $level + 1, $options, ['column' => (string) ($key + 1)]);
                } else {
                    $xml .= $this->valueNode($this->tag((string) $key), $value, $level + 1, $options);
                }
            }

            $xml .= $this->pad($level, $indent) . '</' . $rowTag . '>' . "\n";
        }

        return $xml;
    }

    /** @param array<string,mixed> $options @param array<string,string> $attributes */
    private function valueNode(string $tag, mixed $value, int $level, array $options, array $attributes = []): string
    {
        $indent = (bool) ($options['pretty'] ?? true);
        $tag = $this->tag($tag);
        $attrText = '';
        foreach ($attributes as $name => $attrValue) {
            $attrText .= ' ' . $this->tag($name) . '="' . $this->escape($attrValue) . '"';
        }

        $value = $this->normalizeValue($value);
        if (is_array($value)) {
            $xml = $this->pad($level, $indent) . '<' . $tag . $attrText . '>' . "\n";
            foreach ($value as $key => $item) {
                $childTag = is_int($key) ? 'item' : $this->tag((string) $key);
                $xml .= $this->valueNode($childTag, $item, $level + 1, $options);
            }
            $xml .= $this->pad($level, $indent) . '</' . $tag . '>' . "\n";
            return $xml;
        }

        if ($value === null) {
            return $this->pad($level, $indent) . '<' . $tag . $attrText . '/>' . "\n";
        }

        return $this->pad($level, $indent) . '<' . $tag . $attrText . '>' . $this->escape((string) $value) . '</' . $tag . '>' . "\n";
    }


    /**
     * Encode any XML-safe payload and return the XML string.
     *
     * This is useful for structured workbook outputs where the payload
     * is already prepared by a reader/session and should be returned
     * directly to a controller, API response, or variable without saving.
     *
     * @param array<string,mixed> $options
     */
    public function payloadToString(mixed $payload, array $options = []): string
    {
        $includeDeclaration = (bool) ($options['declaration'] ?? true);
        $root = $this->tag((string) ($options['root'] ?? 'workbook'));
        $xml = $includeDeclaration ? '<?xml version="1.0" encoding="UTF-8"?>' . "\n" : '';

        return $xml . $this->valueNode($root, $payload, 0, $options);
    }

    private function normalizeValue(mixed $value): mixed
    {
        if ($value instanceof CellValue) {
            return $value->displayValue();
        }
        if ($value instanceof \DateTimeInterface) {
            return $value->format('Y-m-d H:i:s');
        }
        if (is_bool($value)) {
            return $value ? 'true' : 'false';
        }
        if (is_array($value)) {
            $out = [];
            foreach ($value as $key => $item) {
                $out[$key] = $this->normalizeValue($item);
            }
            return $out;
        }
        if (is_scalar($value) || $value === null) {
            return $value;
        }

        return (string) $value;
    }

    private function tag(string $name): string
    {
        $name = trim($name);
        $name = preg_replace('/[^A-Za-z0-9_.-]+/', '_', $name) ?: 'item';
        if (!preg_match('/^[A-Za-z_]/', $name)) {
            $name = 'field_' . $name;
        }
        if (str_starts_with(strtolower($name), 'xml')) {
            $name = 'field_' . $name;
        }
        return $name;
    }

    private function escape(string $value): string
    {
        return htmlspecialchars($value, ENT_XML1 | ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
    }

    private function pad(int $level, bool $enabled): string
    {
        return $enabled ? str_repeat('  ', $level) : '';
    }

    private function writeString(string $path, string $contents): void
    {
        AtomicFileWriter::writeString($path, $contents, ErrorCode::XML_WRITE_FAILED);
    }
}
