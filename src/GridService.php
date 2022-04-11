<?php

namespace AnourValar\Office;

class GridService
{
    /**
     * @var string
     */
    protected string $driverClass;

    /**
     * Actions with template before data inserted
     *
     * @var \Closure(GridInterface $gridDriver, array &$headers, iterable &$data, string $leftTopCorner)|null
     */
    protected ?\Closure $hookBefore = null;

    /**
     * Header handler
     *
     * @var \Closure(GridInterface $gridDriver, mixed $header, string|int $key, string $column)|null
     */
    protected ?\Closure $hookHeader = null;

    /**
     * Row data handler
     *
     * @var \Closure(GridInterface $gridDriver, mixed $row, string|int $key)|null
     */
    protected ?\Closure $hookRow = null;

    /**
     * Actions with template after data inserted
     *
     * @var \Closure(GridInterface $gridDriver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns)|null
     */
    protected ?\Closure $hookAfter = null;

    /**
     * @param string $driverClass
     * @return void
     */
    public function __construct(string $driverClass = \AnourValar\Office\Drivers\PhpSpreadsheetDriver::class)
    {
        $this->driverClass = $driverClass;
    }

    /**
     * Generate a document from data (grid)
     *
     * @param array $headers
     * @param iterable|\Closure<iterable> $data
     * @param string $leftTopCorner
     * @return \AnourValar\Office\Generated
     */
    public function generate(array $headers, iterable|\Closure $data, string $leftTopCorner = 'A1'): Generated
    {
        // Get instance of driver
        $driver = new $this->driverClass();
        if (! $driver instanceof \AnourValar\Office\Drivers\GridInterface) {
            throw new \LogicException('Driver must implements GridInterface.');
        }

        // Handle with data
        if ($data instanceof \Closure) {
            $data = $data();
        }

        // Create new document
        $driver->create();

        // Hook: before
        if ($this->hookBefore) {
            ($this->hookBefore)($driver, $headers, $data, $leftTopCorner);
        }

        // Set data
        $driver->setGrid(
            $this->getGenerator($driver, $headers, $data, $leftTopCorner, $headersRange, $dataRange, $totalRange, $columns)()
        );

        // Hook: after
        if ($this->hookAfter) {
            ($this->hookAfter)($driver, $headersRange, $dataRange, $totalRange, $columns);
        }

        return new Generated($driver);
    }

    /**
     * Set hookBefore
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookBefore(?\Closure $closure): self
    {
        $this->hookBefore = $closure;

        return $this;
    }

    /**
     * Set hookHeader
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookHeader(?\Closure $closure): self
    {
        $this->hookHeader = $closure;

        return $this;
    }

    /**
     * Set hookRow
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookRow(?\Closure $closure): self
    {
        $this->hookRow = $closure;

        return $this;
    }

    /**
     * Set hookAfter
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookAfter(?\Closure $closure): self
    {
        $this->hookAfter = $closure;

        return $this;
    }

    /**
     * Set driver class
     *
     * @param string $driverClass
     * @return self
     */
    public function setDriverClass(string $driverClass): self
    {
        $this->driverClass = $driverClass;

        return $this;
    }

    /**
     * @param \AnourValar\Office\Drivers\GridInterface $driver
     * @param array $headers
     * @param iterable $data
     * @param string $leftTopCorner
     * @return \Generator
     */
    protected function getGenerator(
        \AnourValar\Office\Drivers\GridInterface $driver,
        array $headers,
        iterable &$data,
        string $leftTopCorner,
        &$headersRange = null,
        &$dataRange = null,
        &$totalRange = null,
        &$columns = null
    ): \Closure {
        return function () use ($driver, $headers, &$data, $leftTopCorner, &$headersRange, &$dataRange, &$totalRange, &$columns)
        {
            $ltc = preg_split('|([A-Z]+)|', $leftTopCorner, -1, PREG_SPLIT_DELIM_CAPTURE | PREG_SPLIT_NO_EMPTY);

            // left top corner: row
            $headerRow = 1;
            while ($ltc[1] > 1) {
                $ltc[1]--;
                $headerRow++;
                yield [];
            }

            // left top corner: column
            $firstColumn = 'A';
            $indent = [];
            while ($firstColumn < $ltc[0]) {
                $firstColumn++;
                $indent[] = '';
            }

            // Handle with header
            $lastColumn = $firstColumn;
            $isFirst = true;
            $hasHeaders = false;
            foreach ($headers as $key => &$header) {
                if ($isFirst) {
                    $isFirst = false;
                } else {
                    $lastColumn++;
                }

                if ($this->hookHeader) {
                    // Hook: header
                    $header = ($this->hookHeader)($driver, $header, $key, $lastColumn);
                }

                if ($header) {
                    $hasHeaders = true;
                }
            }
            unset($header);

            // First iteration with headers
            if ($hasHeaders) {
                yield array_merge($indent, $headers);
            } else {
                $headerRow--;
            }

            // Iterations with data
            $dataRow = $headerRow;
            $isFirst = ! $hasHeaders;
            foreach ($data as $key => $row) {
                // Hook: row
                if ($this->hookRow) {
                    $row = ($this->hookRow)($driver, $row, $key);
                }

                if (is_null($row)) {
                    continue;
                }

                if ($isFirst) {
                    foreach ($row as $item) {
                        if ($isFirst) {
                            $isFirst = false;
                        } else {
                            $lastColumn++;
                        }
                    }
                }

                yield array_merge($indent, $row);
                $dataRow++;
            }

            // Statistic
            $headersRange = null;
            if ($hasHeaders) {
                $headersRange = sprintf("%s%d:%s%d", $firstColumn, $headerRow, $lastColumn, $headerRow);
            }

            $dataRange = null;
            if ($dataRow != $headerRow) {
                $dataRange = sprintf("%s%d:%s%d", $firstColumn, ($headerRow + 1), $lastColumn, $dataRow);
            }

            $totalRange = null;
            if ($hasHeaders || $dataRow != $headerRow) {
                $totalRange = sprintf("%s%d:%s%d", $firstColumn, ($hasHeaders ? $headerRow : ($headerRow + 1)), $lastColumn, $dataRow);
            }

            $columns = [];
            if ($totalRange) {
                $keys = array_keys($headers);

                while ($firstColumn <= $lastColumn) {
                    if (! $keys) {
                        $columns[] = $firstColumn;
                    } else {
                        $columns[array_shift($keys)] = $firstColumn;
                    }

                    $firstColumn++;
                }
            }
        };
    }
}
