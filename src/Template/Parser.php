<?php

namespace AnourValar\Office\Template;

class Parser
{
    /**
     * Handle with special types of data
     *
     * @param mixed $data
     * @throws \LogicException
     * @return array
     */
    public function canonizeData(mixed $data): array
    {
        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }

        if (! is_array($data)) {
            throw new \LogicException('Data must be an array.');
        }

        return $data;
    }

    /**
     * Get schema for a document
     *
     * @param array $values
     * @param array $data
     * @return array
     */
    public function schema(array $values, array $data): array
    {
        ksort($values); // just in case ;)
        $data = $this->parseData($data);

        // Output contract
        $result = [
            'data' => [],
                /*1 => [
                    'A' => 'foo',
                ],
                2 => [
                    'B' => 'bar',
                ],*/
            'rows' => [
                //['action' => 'add', 'row' => 1],
                //['action' => 'delete', 'row' => 2],
            ],

            'copy_styles' => [
                //['from' => 'B', 'to' => 'C'],
            ],

            'width' => [
                //'C' => 'B'
            ],
        ];

        // Step 0: short path -> full path
        foreach ($values as $row => &$columns) {
            foreach ($columns as $column => &$value) {
                if (! is_string($value)) {
                    continue;
                }

                $value = preg_replace_callback(
                    '#\[(\!\s*|\=\s*)?([a-z\d\.\_]+)\]#i',
                    function ($patterns) use ($data)
                    {
                        if (isset($data[$patterns[2]])) {
                            return $patterns[0];
                        }

                        $result = null;
                        foreach (explode('.', $patterns[2]) as $pattern) {
                            $changed = true;
                            while ($changed) {
                                $changed = false;

                                if ($this->isShortPath($result ? ($result . '.' . $pattern) : $pattern, $data)) {
                                    if ($result) {
                                        $result .= '.';
                                    }

                                    $result .= $pattern;
                                    $changed = true;
                                }

                                if ($this->isShortPath($result . '.0', $data)) {
                                    $result .= '.0';
                                    $changed = true;
                                }

                                if (array_key_exists($result, $data)) {
                                    break 2;
                                }
                            }
                        }

                        if ($result && array_key_exists($result, $data) && !is_array($data[$result])) {
                            return sprintf("[%s%s]", $patterns[1], $result);
                        }

                        return $patterns[0];
                    },
                    $value
                );
            }
            unset($value);
        }
        unset($columns);

        // Step 1: Calculate additional rows & columns, outline markers
        $shift = 0;
        foreach ($values as $row => $columns) {
            $additionRows = 0;
            $additionColumns = 0;
            $additionColumn = null;
            $additionColumnValue = null;

            foreach ($columns as $column => $value) {
                foreach (array_keys($data) as $markerName) {
                    if (! $this->hasMarker($markerName, $value)) {
                        continue;
                    }

                    $qty = 0;
                    $pattern = $markerName;
                    while (array_key_exists($pattern = $this->increment($pattern, true), $data)) {
                         $qty++;
                    }
                    $additionRows = max($additionRows, $qty);

                    $qty = 0;
                    $pattern = $markerName;
                    while (array_key_exists($pattern = $this->increment($pattern, false), $data)) {
                        $qty++;
                    }
                    if ($qty) {
                        $additionColumns = max($additionColumns, $qty);
                        $additionColumn = $column;
                        $additionColumnValue = $value;
                    }
                }
            }

            if ($this->shouldBeDeleted($columns, $data)) {
                $result['rows'][] = ['action' => 'delete', 'row' => ($row + $shift)];
                $shift--;
                continue;
            }

            if (! $additionRows) {
                foreach ($columns as $column => $value) {
                    if (! preg_match('#\[(\!\s*|\=\s*)?[a-z\d\_\.]+\]#i', (string) $value)) {
                        unset($columns[$column]);
                    }
                }

                if (! $columns) {
                    unset($values[$row]);
                    continue;
                }
            }

            $curr = $additionColumn;
            while ($additionColumns) {
                $curr++;
                $additionColumns--;

                $additionColumnValue = $this->increments($additionColumnValue, false);
                $columns[$curr] = $additionColumnValue;
                $result['copy_styles'][] = ['from' => $additionColumn.($row + $shift), 'to' => $curr.($row + $shift)];
                $result['width'][$curr] = $additionColumn;
            }

            $result['data'][$row + $shift] = $columns;
            $originalRow = ($row + $shift);
            while ($additionRows) {
                $shift++;
                $additionRows--;

                foreach ($columns as &$column) {
                    $column = $this->increments((string) $column, true);
                }
                unset($column);

                $result['data'][$row + $shift] = $columns;
                $result['rows'][] = ['action' => 'add', 'row' => ($row + $shift)];

                foreach (array_keys($columns) as $curr) {
                    $result['copy_styles'][] = ['from' => $curr.$originalRow, 'to' => $curr.($row + $shift)];
                }
            }
        }

        // Step 2: Replace markers with data
        foreach ($result['data'] as $row => &$columns) {
            foreach ($columns as $column => &$value) {
                if (is_null($value)) {
                    continue;
                }

                if (is_string($value) && mb_strlen($value)) {
                    $value = preg_replace('#\[(\!\s*|\=\s*)[a-z\d\_\.]+\]#i', '', $value);
                    $value = trim($value);
                }

                foreach ($data as $from => $to) {
                    $value = $this->replace($from, $to, $value);
                }

                if (is_string($value) && mb_strlen($value)) {
                    $value = preg_replace('#\[[a-z\d\_\.]+\]#i', '', $value);
                    $value = trim($value);
                }

                if (is_string($value) && !mb_strlen($value)) {
                    $value = null;
                }
            }
            unset($value);
        }
        unset($columns);

        sort($result['copy_styles']);
        return $result;
    }

    /**
     * @param string $path
     * @param array $markers
     * @return bool
     */
    protected function isShortPath(string $path, array $markers): bool
    {
        if (array_key_exists($path, $markers)) {
            return true;
        }

        if (! str_ends_with($path, '.')) {
            $path .= '.';
        }

        foreach (array_keys($markers) as $marker) {
            if (strpos($marker, $path) === 0) {
                return true;
            }
        }

        return false;
    }

    /**
     * @param array $data
     * @return array
     */
    protected function parseData(array $data): array
    {
        $result = [];

        foreach ($this->dot($data) as $key => $value) {
            if (is_array($value) && !$value) {
                continue;
            }

            $result[$key] = $value;
        }

        return $result;
    }

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
     * @param string $marker
     * @param string $value
     * @return bool
     */
    protected function hasMarker(string $marker, ?string $value): bool
    {
        if (strpos((string) $value, "[$marker]") !== false) {
            return true;
        }

        return false;
    }

    /**
     * @param array $columns
     * @param array $data
     * @return bool
     */
    protected function shouldBeDeleted(array $columns, array $data): bool
    {
        foreach ($columns as $column) {
            if (is_null($column)) {
                continue;
            }

            preg_match_all('#\[\=\s*([a-z\d\.\_]+)\]#i', $column, $patterns);
            foreach (($patterns[1] ?? []) as $marker) {
                if (! empty($data[$marker])) {
                    continue;
                }

                foreach ($data as $key => $value) {
                    if (strpos($key, $marker.'.') === 0 && !empty($value)) {
                        continue 2;
                    }
                }

                return true;
            }

            preg_match_all('#\[\!\s*([a-z\d\.\_]+)\]#i', $column, $patterns);
            foreach (($patterns[1] ?? []) as $marker) {
                if (! empty($data[$marker])) {
                    return true;
                }

                foreach ($data as $key => $value) {
                    if (strpos($key, $marker.'.') === 0 && !empty($value)) {
                        return true;
                    }
                }

                return false;
            }
        }

        return false;
    }

    /**
     * @param string $markerName
     * @param bool $first
     * @return string|null
     */
    protected function increment(string $markerName, bool $first): ?string
    {
        $markerName = explode('.', $markerName);
        if (! $first) {
            $markerName = array_reverse($markerName);
        }

        if (! $first) {
            $qty = 0;
            foreach ($markerName as $item) {
                if (is_numeric($item)) {
                    $qty++;
                }
            }

            if ($qty < 2) {
                return null;
            }
        }

        foreach ($markerName as &$item) {
            if (is_numeric($item)) {
                $item++;

                if (! $first) {
                    $markerName = array_reverse($markerName);
                }

                return implode('.', $markerName);
            }
        }
        unset($item);

        return null;
    }

    /**
     * @param string $value
     * @param bool $first
     * @return string|null
     */
    protected function increments(string $value, bool $first): ?string
    {
        return preg_replace_callback(
            '#\[([a-z\d\.\_]+)\]#i',
            function ($patterns) use ($first)
            {
                $patterns[1] = $this->increment($patterns[1], $first);
                if ($patterns[1]) {
                    return '[' . $patterns[1] . ']';
                }

                return $patterns[0];
            },
            $value
        );
    }

    /**
     * @param string $from
     * @param mixed $to
     * @param mixed $value
     * @throws \LogicException
     * @return mixed
     */
    protected function replace(string $from, mixed $to, $value): mixed
    {
        if (! is_string($value)) {
            return $value;
        }

        $from = "[$from]";

        if ($value == $from) {
            return $to;
        }

        if (strpos($value, $from) === false) {
            return $value;
        }

        if ($to instanceof \Closure) {
            throw new \LogicException('Parameter cannot be used as part of cell\'s value if it\'s a closure.');
        }

        return str_replace($from, (string) $to, $value);
    }
}
