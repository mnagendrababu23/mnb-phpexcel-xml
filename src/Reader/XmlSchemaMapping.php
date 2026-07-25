<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Reader;

use Mnb\PHPExcel\Support\MnbExcelException;

/**
 * Lightweight streaming XML schema map. It intentionally supports a safe,
 * deterministic subset of XPath: child paths (`customer/name`), row
 * attributes (`@id`), and row text (`text()`).
 */
final class XmlSchemaMapping
{
    /** @param array<string,array<string,mixed>> $columns */
    private function __construct(
        private readonly string $sheetTag,
        private readonly string $sheetNameAttribute,
        private readonly string $rowTag,
        private readonly array $columns,
    ) {
    }

    /** @param array<string,mixed>|self|null $schema */
    public static function from(array|self|null $schema, array $fallbackOptions = []): ?self
    {
        if ($schema instanceof self) {
            return $schema;
        }
        if ($schema === null || $schema === []) {
            return null;
        }

        $columns = [];
        foreach ((array) ($schema['columns'] ?? []) as $target => $definition) {
            if (is_int($target)) {
                $target = is_string($definition) ? $definition : (string) ($definition['target'] ?? 'column_' . ($target + 1));
            }
            if (is_string($definition)) {
                $definition = ['source' => $definition];
            }
            if (!is_array($definition)) {
                throw new MnbExcelException('Each XML schema column must be a source string or configuration array.');
            }
            $columns[(string) $target] = $definition + ['source' => (string) $target, 'type' => 'mixed'];
        }

        return new self(
            (string) ($schema['sheet_tag'] ?? $fallbackOptions['sheet_tag'] ?? 'sheet'),
            (string) ($schema['sheet_name_attribute'] ?? 'name'),
            (string) ($schema['row_tag'] ?? $fallbackOptions['row_tag'] ?? $fallbackOptions['row'] ?? 'row'),
            $columns,
        );
    }

    public function sheetTag(): string
    {
        return $this->sheetTag;
    }

    public function sheetNameAttribute(): string
    {
        return $this->sheetNameAttribute;
    }

    public function rowTag(): string
    {
        return $this->rowTag;
    }

    /**
     * @param array<int|string,mixed> $row
     * @param array<string,string> $attributes
     * @return array<int|string,mixed>
     */
    public function map(array $row, array $attributes = [], string $rowText = ''): array
    {
        if ($this->columns === []) {
            return $row;
        }

        $mapped = [];
        foreach ($this->columns as $target => $definition) {
            $source = trim((string) ($definition['source'] ?? $target));
            $found = true;
            if (str_starts_with($source, '@')) {
                $attribute = substr($source, 1);
                $found = array_key_exists($attribute, $attributes);
                $value = $attributes[$attribute] ?? null;
            } elseif ($source === 'text()' || $source === '#text') {
                $value = $rowText;
            } else {
                [$found, $value] = $this->valueAtPath($row, $source);
            }

            if (!$found) {
                if (array_key_exists('default', $definition)) {
                    $value = $definition['default'];
                } elseif ((bool) ($definition['required'] ?? false)) {
                    throw new MnbExcelException('Required XML schema source is missing: ' . $source);
                } else {
                    $value = null;
                }
            }

            $mapped[$target] = $this->cast($value, (string) ($definition['type'] ?? 'mixed'), $definition);
        }
        return $mapped;
    }

    /** @param array<int|string,mixed> $row @return array{bool,mixed} */
    private function valueAtPath(array $row, string $path): array
    {
        $parts = array_values(array_filter(preg_split('~[/.]+~', $path) ?: [], static fn(string $part): bool => $part !== ''));
        if ($parts === []) {
            return [false, null];
        }
        $value = $row;
        foreach ($parts as $part) {
            if (!is_array($value) || !array_key_exists($part, $value)) {
                return [false, null];
            }
            $value = $value[$part];
        }
        return [true, $value];
    }

    /** @param array<string,mixed> $definition */
    private function cast(mixed $value, string $type, array $definition): mixed
    {
        if ($value === null) {
            return null;
        }
        return match (strtolower($type)) {
            '', 'mixed', 'auto' => $value,
            'string', 'text' => (string) $value,
            'int', 'integer' => (int) $value,
            'float', 'double', 'number', 'decimal' => (float) $value,
            'bool', 'boolean' => filter_var($value, FILTER_VALIDATE_BOOL, FILTER_NULL_ON_FAILURE) ?? false,
            'json' => is_string($value) ? json_decode($value, true, 512, JSON_THROW_ON_ERROR) : $value,
            'date', 'datetime' => $this->castDate($value, $definition),
            default => throw new MnbExcelException('Unsupported XML schema type: ' . $type),
        };
    }

    /** @param array<string,mixed> $definition */
    private function castDate(mixed $value, array $definition): mixed
    {
        try {
            $date = new \DateTimeImmutable((string) $value);
        } catch (\Throwable) {
            return $value;
        }
        if ((bool) ($definition['return_datetime'] ?? false)) {
            return $date;
        }
        return $date->format((string) ($definition['format'] ?? 'Y-m-d H:i:s'));
    }
}
