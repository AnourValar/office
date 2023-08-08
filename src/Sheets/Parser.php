<?php

namespace AnourValar\Office\Sheets;

class Parser
{
    use \AnourValar\Office\Traits\Parser;

    /**
     * Handle with special types of data
     *
     * @param mixed $data
     * @return array
     */
    public function canonizeData(mixed $data): array
    {
        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }

        return $data;
    }

    /**
     * Get schema for a document
     *
     * @param array $values
     * @param array $data
     * @param array $mergeCells
     * @return \AnourValar\Office\Sheets\SchemaMapper
     */
    public function schema(array $values, array $data, array $mergeCells): SchemaMapper
    {
        $schema = new SchemaMapper();

        // Step 0: Parse input arguments to a canon format
        $values = $this->parseValues($values, $lastColumn);
        $data = $this->parseData($data);
        $mergeCells = $this->parseMergeCells($mergeCells);

        // Step 1: Short path -> full path
        $this->canonizeMarkers($values, $data);

        // Step 2: Calculate additional rows & columns, redundant data
        $dataSchema = $this->calculateDataSchema($values, $data, $mergeCells, $schema, $lastColumn);

        // Step 3: Shift formulas
        $this->shiftFormulas($dataSchema, $schema, $mergeCells);

        // Step 4: Replace markers with data
        $this->replaceMarkers($dataSchema, $data, $schema);

        return $schema;
    }

    /**
     * @param array $values
     * @param mixed $lastColumn
     * @return array
     */
    protected function parseValues(array $values, &$lastColumn): array
    {
        $lastColumn = 'A';
        foreach ($values as $row => &$columns) {
            $currLastColumn = array_key_last($columns);
            if ($this->isColumnLE($lastColumn, $currLastColumn)) {
                $lastColumn = $currLastColumn;
            }

            $columns = array_filter($columns, fn ($item) => $item !== null && $item !== '');
        }
        unset($columns);

        return $values;
    }

    /**
     * @param array $data
     * @return array
     */
    protected function parseData(array &$data): array
    {
        $result = [];

        foreach ($this->dot($data) as $key => $value) {
            if (is_array($value) && ! $value) {
                continue;
            }

            $result[$key] = $value;
        }

        return $result;
    }

    /**
     * @param array $mergeCells
     * @return array
     */
    protected function parseMergeCells(array $mergeCells): array
    {
        foreach ($mergeCells as &$item) {
            $item = explode(':', $item);

            $item[0] = preg_split('#([A-Z]+)([\d]+)#S', $item[0], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
            $item[1] = preg_split('#([A-Z]+)([\d]+)#S', $item[1], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
        }
        unset($item);

        return $mergeCells;
    }

    /**
     * @param array $values
     * @param array $data
     * @return void
     */
    protected function canonizeMarkers(array &$values, array &$data): void
    {
        foreach ($values as &$columns) {
            foreach ($columns as &$value) {
                if (! is_string($value)) {
                    continue;
                }

                $value = preg_replace_callback(
                    '#\[(\$?\!\s*|\$?\=\s*)?([a-z\d\.\_\*]+)\]#iS',
                    function ($patterns) use ($data) {
                        if (array_key_exists($patterns[2], $data)) {
                            return $patterns[0];
                        }

                        $result = null;
                        foreach (explode('.', $patterns[2]) as $pattern) {
                            $changed = true;
                            $prevResult = $result;
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
                            }

                            if ($result === $prevResult && ($pattern != '*' || mb_substr($result, -2) != '.0')) {
                                return $patterns[0];
                            }
                        }

                        if ($result && array_key_exists($result, $data) && ! is_array($data[$result])) {
                            $result = preg_replace('#\.0(\.|$)#S', '.*$1', $result);
                            return sprintf('[%s%s]', $patterns[1], $result);
                        }

                        return $patterns[0];
                    },
                    $value
                );
            }
            unset($value);
        }
        unset($columns);
    }

    /**
     * @param array $values
     * @param array $data
     * @param array $mergeCells
     * @param \AnourValar\Office\Sheets\SchemaMapper $schema
     * @param string $lastColumn
     * @return array
     */
    protected function calculateDataSchema(
        array &$values,
        array &$data,
        array &$mergeCells,
        SchemaMapper &$schema,
        string $lastColumn
    ): array {
        $dataSchema = [];
        $shift = 0;
        $step = 0;
        $stepLeft = 0;
        $stepOrigin = 0;
        $stepRows = 0;

        // fill in missing rows
        $prevRow = 0;
        foreach (array_keys($values) as $row) {
            $diff = ($row - $prevRow);
            while ($diff > 1) {
                $values[$row - $diff + 1] = [];
                $diff--;
            }

            $prevRow = $row;
        }
        ksort($values);

        foreach ($values as $row => $columns) {
            $maxMergeY = 0;

            $additionRows = 0;
            $additionColumns = 0;
            $additionColumn = null;

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
                    }
                }

                if (is_string($value)) {
                    $columns[$column] = preg_replace('#\.\*(\.|\])#S', '.0$1', $value);
                }
            }


            if (! $stepRows && $this->shouldBeDeleted($columns, $data)) {
                $this->deleteRow($schema, $mergeCells, $row + $shift);
                $shift--;
                continue;
            }

            if ($stepRows) {
                $additionRows = $stepRows;
            }
            $currAdditionRows = $additionRows;

            if (! $additionRows) {
                foreach ($columns as $column => $value) {
                    if (
                        ! preg_match('#\[(\$?\!\s*|\$?\=\s*)?[a-z][a-z\d\_\.]+\]#iS', (string) $value)
                        && ! preg_match('#^\=[A-Z][A-Z\.\d]#', (string) $value)
                    ) {
                        unset($columns[$column]);
                    }
                }

                if (! $columns) {
                    continue;
                }
            }

            $curr = $additionColumn;
            $additionColumnValue = ($columns[$curr] ?? null);

            $mergeMapX = [];
            foreach ($mergeCells as $item) {
                if ($additionColumn.($row + $shift) == $item[0][0].$item[0][1] && $item[0][1] == $item[1][1]) {
                    while ($this->isColumnLE($item[0][0], $item[1][0]) && $item[0][0] != $item[1][0]) {
                        $item[0][0]++;
                        $mergeMapX[] = $item[0][0];
                    }
                }
            }

            foreach ($mergeMapX as $mergeItemX) {
                $curr++;
            }

            while ($additionColumns) {
                $curr++;
                $additionColumns--;

                $additionColumnValue = $this->increments($additionColumnValue, false);
                $columns[$curr] = $additionColumnValue;
                $schema->copyStyle($additionColumn.($row + $shift), $curr.($row + $shift));
                $schema->copyWidth($additionColumn, $curr);

                if ($mergeMapX) {
                    $originalCurr = $curr;
                    foreach ($mergeMapX as $mergeItemX) {
                        $curr++;

                        $schema->copyWidth($mergeItemX, $curr);
                    }

                    $schema->mergeCells(sprintf('%s%s:%s%s', $originalCurr, ($row + $shift), $curr, ($row + $shift)));
                    $mergeCells[] = [ [$originalCurr, ($row + $shift)], [$curr, ($row + $shift)] ]; // fill in
                }
            }

            $dataSchema[$row + $shift] = $columns;
            $originalRow = ($row + $shift);

            if ($additionRows) {
                $firstColumn = 'A';
                while ($this->isColumnLE($firstColumn, $lastColumn)) {
                    if (! isset($columns[$firstColumn]) && ! $this->insideMerge($firstColumn, $originalRow, $mergeCells)) {
                        $columns[$firstColumn] = null;
                    }

                    $firstColumn++;
                }
                uksort($columns, fn ($a, $b) => $this->isColumnLE($a, $b) ? -1 : 1);

                foreach ($columns as $currKey => $currValue) {
                    $hasMarker = preg_match('#\[([a-z][a-z\d\.\_]+)\]#iS', (string) $currValue);

                    foreach ($mergeCells as $item) {
                        if ($currKey.$originalRow == $item[0][0].$item[0][1] && $item[0][1] != $item[1][1]) {
                            if (! $hasMarker) {
                                unset($columns[$currKey]);
                                continue;
                            }

                            $maxMergeY = max($maxMergeY, ($item[1][1] - $item[0][1]));
                        }
                    }
                }

                $shift += $maxMergeY;
            }

            while ($additionRows) {
                $shift += $step;
                $shift++;
                $additionRows--;

                foreach ($columns as &$column) {
                    if (is_string($column)) {
                        $column = $this->increments($column, true);
                    }
                }
                unset($column);

                $dataSchema[$row + $shift] = array_filter($columns, fn ($item) => $item !== null && $item !== '');
                if (! $step) {
                    $this->addRow($schema, $mergeCells, $row + $shift);
                }

                foreach (array_keys($columns) as $curr) {
                    $schema->copyStyle($curr.$originalRow, $curr.($row + $shift));

                    foreach ($mergeCells as $item) {
                        if ($curr.$originalRow == $item[0][0].$item[0][1]) {
                            $diff = $item[1][1] - $item[0][1];
                            $schema->mergeCells(
                                sprintf('%s%s:%s%s', $item[0][0], ($row + $shift), $item[1][0], ($row + $shift + $diff))
                            );
                        }
                    }
                }

                $iterate = $maxMergeY;
                while ($iterate) {
                    $shift++;
                    $this->addRow($schema, $mergeCells, $row + $shift);
                    $iterate--;
                }
            }

            if ($stepLeft) {
                $stepLeft--;
            }
            if (! $stepLeft) {
                $step = 0;
                $stepRows = 0;
            } else {
                $stepOrigin--;
                $shift -= $stepOrigin - $stepLeft;
            }

            if ($maxMergeY) {
                $stepRows = $currAdditionRows;
                $step = $maxMergeY;
                $stepLeft = $step;
                $stepOrigin = (($maxMergeY + 1) * ($currAdditionRows)) + $step;
                $shift -= $stepOrigin;
            }
        }
        unset($values);

        return $dataSchema;
    }

    /**
     * @param array $values
     * @param \AnourValar\Office\Sheets\SchemaMapper $schema
     * @param array $mergeCells
     * @return void
     */
    protected function shiftFormulas(array &$values, SchemaMapper &$schema, array &$mergeCells): void
    {
        // Prepares
        $ranges = [];

        $map = [];
        foreach ($values as $row => $columns) {
            foreach ($columns as $column => $value) {
                if (preg_match('#^\=[A-Z][A-Z\.\d]#', (string) $value)) {
                    $map[$row][$column] = $value;
                }
            }
        }

        // "Outside" shifts
        foreach ($schema->getOriginal()['rows'] as $action) {
            foreach ($map as $row => $columns) {
                foreach ($columns as $column => $value) {
                    $map[$row][$column] = $values[$row][$column] = preg_replace_callback(
                        '#([A-Z]+)([\d]+)#S',
                        function ($patterns) use ($action) {
                            if ($action['action'] == 'add') {

                                if ($patterns[2] >= $action['row']) {
                                    return $patterns[1] . ++$patterns[2];
                                }

                            } else {

                                if ($patterns[2] > $action['row']) {
                                    return $patterns[1] . --$patterns[2];
                                }

                            }

                            return $patterns[0];
                        },
                        $value
                    );
                }
            }
        }

        // "Inside" shifts
        $prev = 0;
        $prevAction = null;
        foreach ($schema->getOriginal()['rows'] as $action) {
            if ($prevAction && $prevAction['row'] + 1 == $action['row'] && $action['action'] == 'add' && $prevAction['action'] == 'add') {
                $prev++;
            } else {
                if ($prevAction && $prevAction['action'] == 'add') {
                    $ranges[] = ['from' => ($prevAction['row'] - $prev - 1), 'to' => ($prevAction['row'])];
                }

                $prev = 0;
            }

            foreach ($map as $row => $columns) {
                foreach ($columns as $column => $value) {
                    $map[$row][$column] = $values[$row][$column] = preg_replace_callback(
                        '#([A-Z]+)([\d]+)#S',
                        function ($patterns) use ($action, $row, $prev) {
                            if ($action['action'] == 'add') {

                                if ($action['row'] == $row && ($row - $prev) == ($patterns[2] + 1)) {
                                    return $patterns[1] . (++$patterns[2] + $prev);
                                }

                            }

                            return $patterns[0];
                        },
                        $value
                    );
                }
            }

            $prevAction = $action;
        }
        if ($prev || ($prevAction && $prevAction['action'] == 'add')) {
            $ranges[] = ['from' => ($prevAction['row'] - $prev - 1), 'to' => ($prevAction['row'])];
        }

        // Dynamic table ranges
        foreach ($ranges as $key => $value) {
            foreach ($mergeCells as $merge) {
                if (
                    $value['from'] >= $merge[0][1]
                    && $value['from'] <= $merge[1][1]
                    && $merge[0][1] != $merge[1][1]
                    && preg_match('#\[([a-z][a-z\d\.\_]+)\]#iS', ($values[$merge[0][1]][$merge[0][0]] ?? ''))
                ) {
                    unset($ranges[$key]);
                }
            }
        }

        foreach ($map as $row => $columns) {
            foreach ($columns as $column => $value) {
                $map[$row][$column] = $values[$row][$column] = preg_replace_callback(
                    '#([A-Z]+)([\d]+)\:([A-Z]+)([\d]+)#S',
                    function ($patterns) use ($ranges) {
                        if ($patterns[2] == $patterns[4]) {
                            foreach ($ranges as $range) {
                                if ($patterns[2] == $range['from']) {
                                    return sprintf('%s%d:%s%d', $patterns[1], $patterns[2], $patterns[3], $range['to']);
                                }
                            }
                        }

                        return $patterns[0];
                    },
                    $value
                );
            }
        }
    }

    /**
     * @param array $dataSchema
     * @param array $data
     * @param \AnourValar\Office\Sheets\SchemaMapper $schema
     * @return void
     */
    protected function replaceMarkers(array &$dataSchema, array &$data, SchemaMapper &$schema): void
    {
        $canonizeKeys = ['scalar' => [], 'closure' => []];
        $canonizeValues = ['scalar' => [], 'closure' => []];
        foreach ($data as $from => $to) {
            if (is_scalar($to) || is_null($to)) {
                $canonizeKeys['scalar'][] = "[$from]";
                $canonizeValues['scalar'][] = $to;
            } else {
                $canonizeKeys['closure'][] = "[$from]";
                $canonizeValues['closure'][] = $to;
            }
        }

        foreach ($dataSchema as $row => $columns) {
            ksort($columns);

            foreach ($columns as $column => $value) {
                if (is_string($value) && $this->shouldBeDeleted([$value], $data, '$')) {
                    $value = null;
                }

                if (is_string($value) && mb_strlen($value)) {
                    $value = preg_replace('#\[(\$?\!\s*|\$?\=\s*)[a-z][a-z\d\_\.]+\]#iS', '', $value);
                    $value = trim($value);

                    if (($key = array_search($value, $canonizeKeys['scalar'])) !== false) { // type (cast) support
                        $value = $canonizeValues['scalar'][$key];
                    } elseif (($key = array_search($value, $canonizeKeys['closure'])) !== false) {
                        $value = $canonizeValues['closure'][$key];
                    } else {
                        $value = str_replace($canonizeKeys['scalar'], $canonizeValues['scalar'], $value);
                    }
                }

                if (is_string($value) && mb_strlen($value)) {
                    $value = preg_replace('#\[[a-z][a-z\d\_\.]+\]#iS', '', $value);
                    $value = trim($value);
                }

                if (is_string($value) && ! mb_strlen($value)) {
                    $value = null;
                }

                $schema->addData($row, $column, $value);
            }
        }
    }

    /**
     * @param string $path
     * @param array $markers
     * @return bool
     */
    private function isShortPath(string $path, array $markers): bool
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
     * @param string $marker
     * @param string $value
     * @return bool
     */
    private function hasMarker(string $marker, ?string $value): bool
    {
        $value = preg_replace('#\.\*(\.|$|)#S', '.0$1', (string) $value, -1, $count);
        if (! $count) {
            return false;
        }

        if (strpos((string) $value, "[$marker]") !== false) {
            return true;
        }

        return false;
    }

    /**
     * @param array $columns
     * @param array $data
     * @param string $prefix
     * @return bool
     */
    private function shouldBeDeleted(array $columns, array &$data, string $prefix = ''): bool
    {
        $prefix = preg_quote($prefix);

        foreach ($columns as $column) {
            if (is_null($column)) {
                continue;
            }

            preg_match_all("#\[{$prefix}\=\s*([a-z\d\.\_]+)\]#i", $column, $patterns);
            foreach (($patterns[1] ?? []) as $marker) {
                if (! empty($data[$marker])) {
                    continue;
                }

                foreach ($data as $key => $value) {
                    if (strpos($key, $marker.'.') === 0 && ! empty($value)) {
                        continue 2;
                    }
                }

                return true;
            }

            preg_match_all("#\[{$prefix}\!\s*([a-z\d\.\_]+)\]#i", $column, $patterns);
            foreach (($patterns[1] ?? []) as $marker) {
                if (! empty($data[$marker])) {
                    return true;
                }

                foreach ($data as $key => $value) {
                    if (strpos($key, $marker.'.') === 0 && ! empty($value)) {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    /**
     * @param string $markerName
     * @param bool $first
     * @param int $shift
     * @return string|null
     */
    private function increment(string $markerName, bool $first, int $shift = 1): ?string
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
                $item += $shift;

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
     * @param int $shift
     * @return string|null
     */
    private function increments(string $value, bool $first, int $shift = 1): ?string
    {
        return preg_replace_callback(
            '#\[(\$?\!\s*|\$?\=\s*)?([a-z][a-z\d\.\_]+)\]#iS',
            function ($patterns) use ($first, $shift) {
                $patterns[2] = $this->increment($patterns[2], $first, $shift);
                if ($patterns[2]) {
                    return '[' . $patterns[1] . $patterns[2] . ']';
                }

                return $patterns[0];
            },
            $value
        );
    }

    /**
     * @param \AnourValar\Office\Sheets\SchemaMapper $schema
     * @param array $mergeCells
     * @param int $row
     * @return void
     */
    private function addRow(SchemaMapper &$schema, array &$mergeCells, int $row): void
    {
        $schema->addRow($row);

        foreach ($mergeCells as &$mergeCell) {
            if ($mergeCell[0][1] >= $row) {
                $mergeCell[0][1]++;
            }

            if ($mergeCell[1][1] >= $row) {
                $mergeCell[1][1]++;
            }
        }
        unset($mergeCell);
    }

    /**
     * @param \AnourValar\Office\Sheets\SchemaMapper $schema
     * @param array $mergeCells
     * @param int $row
     * @return void
     */
    private function deleteRow(SchemaMapper &$schema, array &$mergeCells, int $row): void
    {
        $schema->deleteRow($row);

        foreach ($mergeCells as &$mergeCell) {
            if ($mergeCell[0][1] >= $row) {
                $mergeCell[0][1]--;
            }

            if ($mergeCell[1][1] >= $row) {
                $mergeCell[1][1]--;
            }
        }
        unset($mergeCell);
    }

    /**
     * @param string $column
     * @param int $row
     * @param array $mergeCells
     * @return bool
     */
    private function insideMerge(string $column, int $row, array &$mergeCells): bool
    {
        foreach ($mergeCells as $item) {
            if (
                $this->isColumnLE($item[0][0], $column)
                && $this->isColumnGE($item[1][0], $column)
                && $item[0][1] <= $row
                && $item[1][1] >= $row
                && ($item[0][0] != $column || $item[0][1] != $row)
            ) {
                return true;
            }
        }

        return false;
    }
}
