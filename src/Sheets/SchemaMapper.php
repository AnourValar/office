<?php

namespace AnourValar\Office\Sheets;

class SchemaMapper
{
    /**
     * @var array
     */
    protected array $payload = [
        'data' => [], //[ 1 => ['A' => 'foo'], '2' => ['B' => 'bar'] ]

        'rows' => [], //[ ['action' => 'add', 'row' => 1, 'qty' => 1], ['action' => 'delete', 'row' => 2, 'qty' => 1] ]

        'copy_style' => [], //[ ['from' => 'A1', 'to' => 'A2'] ]

        'merge_cells' => [], //[ 'A1:B1', 'C1:D1']

        'copy_width' => [], //[ ['from' => 'B', 'to' => 'C'] ]
    ];

    /**
     * @return array
     */
    public function toArray(): array
    {
        ksort($this->payload['data']);

        $this->normalizeRows($this->payload['rows']);

        $this->normalizeCells($this->payload['copy_style']);

        sort($this->payload['copy_width']);

        return $this->payload;
    }

    /**
     * @return array
     */
    public function getOriginal(): array
    {
        return $this->payload;
    }

    /**
     * @param int $row
     * @param string $column
     * @param mixed $value
     * @return self
     */
    public function addData(int $row, string $column, mixed $value): self
    {
        $this->payload['data'][$row][$column] = $value;

        return $this;
    }

    /**
     * @param int $rowBefore
     * @return self
     */
    public function addRow(int $rowBefore): self
    {
        $this->payload['rows'][] = ['action' => 'add', 'row' => $rowBefore, 'qty' => 1];

        return $this;
    }

    /**
     * @param int $row
     * @return self
     */
    public function deleteRow(int $row): self
    {
        $this->payload['rows'][] = ['action' => 'delete', 'row' => $row, 'qty' => 1];

        return $this;
    }

    /**
     * @param string $from
     * @param string $to
     * @return self
     */
    public function copyStyle(string $from, string $to): self
    {
        $this->payload['copy_style'][$from.$to] = ['from' => $from, 'to' => $to];

        return $this;
    }

    /**
     * @param string $ceilRange
     * @return self
     */
    public function mergeCells(string $ceilRange): self
    {
        $this->payload['merge_cells'][] = $ceilRange;

        return $this;
    }

    /**
     * @param string $from
     * @param string $to
     * @return self
     */
    public function copyWidth(string $from, string $to): self
    {
        $this->payload['copy_width'][$from . $to] = ['from' => $from, 'to' => $to];

        return $this;
    }

    /**
     * @param array $rows
     * @return void
     */
    protected function normalizeRows(array &$rows): void
    {
        $optimizedRows = [];

        $curr = [];
        foreach ($rows as $item) {
            if ($curr) {
                if ($item['action'] == 'add' && $curr['action'] == $item['action'] && $item['row'] == ($curr['row'] + $curr['qty'])) {
                    $curr['qty']++;
                } else {
                    $optimizedRows[] = ['action' => $curr['action'], 'row' => $curr['row'], 'qty' => $curr['qty']];
                    $curr = null;
                }
            }

            if (! $curr) {
                $curr = ['action' => $item['action'], 'row' => $item['row'], 'qty' => $item['qty']];
            }
        }

        if ($curr) {
            $optimizedRows[] = ['action' => $curr['action'], 'row' => $curr['row'], 'qty' => $curr['qty']];
        }

        $rows = $optimizedRows;
    }

    /**
     * @param array $data
     * @return void
     */
    protected function normalizeCells(array &$data): void
    {
        ksort($data, SORT_NATURAL);
        $data = array_values($data);

        $optimizedData = [];

        $curr = [];
        foreach ($data as $item) {
            $expect = preg_split('#([A-Z]+)([\d]+)#S', $item['to'], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
            $expect[1]++;
            $expect = implode('', $expect);

            if ($curr) {
                if ($item['from'] == $curr['from'] && $item['to'] == $curr['expect']) {
                    $curr['append'] = ':' . $curr['expect'];
                    $curr['expect'] = $expect;
                } else {
                    $optimizedData[] = ['from' => $curr['from'], 'to' => $curr['to'] . $curr['append']];
                    $curr = null;
                }
            }

            if (! $curr) {
                $curr = ['from' => $item['from'], 'to' => $item['to'], 'append' => '', 'expect' => $expect];
            }
        }

        if ($curr) {
            $optimizedData[] = ['from' => $curr['from'], 'to' => $curr['to'] . $curr['append']];
        }

        $data = $optimizedData;
    }
}
