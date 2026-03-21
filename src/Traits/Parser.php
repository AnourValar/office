<?php

namespace AnourValar\Office\Traits;

trait Parser
{
    /**
     * @param array $data
     * @param string $prefix
     * @return array
     */
    protected function dot(array $data, string $prefix = ''): array
    {
        $result = [];

        foreach ($data as $key => $value) {
            if (is_array($value)) {
                $result = array_replace($result, $this->dot($value, $prefix.$key.'.'));
            } else {
                $result[$prefix.$key] = $value;
            }
        }

        return $result;
    }

    /**
     * @param string $compareColumn
     * @param string $referenceColumn
     * @return bool
     */
    protected function isColumnLE(string $compareColumn, string $referenceColumn): bool
    {
        $compareLength = strlen($compareColumn);
        $referenceLength = strlen($referenceColumn);

        if ($compareLength < $referenceLength) {
            return true;
        }

        if ($compareLength > $referenceLength) {
            return false;
        }

        return $compareColumn <= $referenceColumn;
    }

    /**
     * @param string $compareColumn
     * @param string $referenceColumn
     * @return bool
     */
    protected function isColumnGE(string $compareColumn, string $referenceColumn): bool
    {
        $compareLength = strlen($compareColumn);
        $referenceLength = strlen($referenceColumn);

        if ($compareLength > $referenceLength) {
            return true;
        }

        if ($compareLength < $referenceLength) {
            return false;
        }

        return $compareColumn >= $referenceColumn;
    }

    /**
     * Polyfill
     *
     * @param string $value
     * @return string
     */
    protected function strIncrement(string $value): string
    {
        if (PHP_VERSION_ID >= 80300) {
            return str_increment($value);
        }

        $value++; // @phpstan-ignore-line
        return $value;
    }
}
