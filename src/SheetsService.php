<?php

namespace AnourValar\Office;

use AnourValar\Office\Drivers\SheetsInterface;

class SheetsService
{
    /**
     * @var \AnourValar\Office\Drivers\SheetsInterface
     */
    protected \AnourValar\Office\Drivers\SheetsInterface $driver;

    /**
     * @var \AnourValar\Office\Sheets\Parser
     */
    protected \AnourValar\Office\Sheets\Parser $parser;

    /**
     * Handle template's loading
     *
     * @var \Closure(SheetsInterface $driver, string $templateFile, Format $templateFormat): SheetsInterface
     */
    protected ?\Closure $hookLoad = null;

    /**
     * Actions with template before data inserted
     *
     * @var \Closure(SheetsInterface $driver, array &$data)
     */
    protected ?\Closure $hookBefore = null;

    /**
     * Cell's value handler (on set)
     *
     * @var \Closure(SheetsInterface $driver, string $column, int $row, mixed $value, int $sheetIndex)
     */
    protected ?\Closure $hookValue = null;

    /**
     * Actions with template after data inserted
     *
     * @var \Closure(SheetsInterface $driver)
     */
    protected ?\Closure $hookAfter = null;

    /**
     * @param \AnourValar\Office\Drivers\SheetsInterface $driver
     * @param \AnourValar\Office\Sheets\Parser $parser
     * @return void
     */
    public function __construct(
        SheetsInterface $driver = new \AnourValar\Office\Drivers\PhpSpreadsheetDriver(),
        \AnourValar\Office\Sheets\Parser $parser = new \AnourValar\Office\Sheets\Parser()
    ) {
        $this->driver = $driver;
        $this->parser = $parser;
    }

    /**
     * Generate a document from the template (sheets)
     *
     * @param string|\Stringable $templateFile
     * @param mixed $data
     * @param bool $autoCellFormat
     * @return \AnourValar\Office\Generated
     */
    public function generate(string|\Stringable $templateFile, mixed $data, bool $autoCellFormat = false): Generated
    {
        // Handle with input data
        $data = $this->parser->canonizeData($data);

        // Open the template
        $templateFormat = Format::tryFrom(mb_strtolower(pathinfo($templateFile, PATHINFO_EXTENSION))) ?? Format::Xlsx;

        if ($this->hookLoad) {
            $driver = ($this->hookLoad)($this->driver, $templateFile, $templateFormat);
            if ($driver instanceof Generated) {
                $driver = $driver->driver;
            }
        } else {
            $driver = $this->driver->load($templateFile, $templateFormat);
        }

        // Hook: before
        if ($this->hookBefore) {
            ($this->hookBefore)($driver, $data);
        }

        // Handle sheets
        $count = $driver->getSheetCount();
        for ($sheetIndex = 0; $sheetIndex < $count; $sheetIndex++) {
            $driver->setSheet($sheetIndex);

            $this->handleSheet($driver, $data, $sheetIndex, $autoCellFormat);
        }

        // Hook: after
        if ($this->hookAfter) {
            ($this->hookAfter)($driver);
        }

        // Return
        return new Generated($driver);
    }

    /**
     * Set hookLoad
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookLoad(?\Closure $closure): self
    {
        $this->hookLoad = $closure;

        return $this;
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
     * Set hookValue
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookValue(?\Closure $closure): self
    {
        $this->hookValue = $closure;

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
     * @param \AnourValar\Office\Drivers\SheetsInterface $driver
     * @param array $data
     * @param int $sheetIndex
     * @throws \LogicException
     * @return void
     */
    protected function handleSheet(SheetsInterface &$driver, array &$data, int $sheetIndex, bool $autoCellFormat): void
    {
        // Get schema of the document
        $schema = $this->parser->schema($driver->getValues(null), $data, $driver->getMergeCells())->toArray();

        // Rows
        foreach ($schema['rows'] as $row) {
            if ($row['action'] == 'add') {
                $driver->addRow($row['row'], $row['qty']);
            } elseif ($row['action'] == 'delete') {
                $driver->deleteRow($row['row'], $row['qty']);
            } else {
                throw new \LogicException('Incorrect usage.');
            }
        }

        // Copy style & cell format
        foreach ($schema['copy_style'] as $item) {
            $driver->copyStyle($item['from'], $item['to']);

            if (! $autoCellFormat) {
                $driver->copyCellFormat($item['from'], $item['to']);
            }
        }

        // Merge cells
        foreach ($schema['merge_cells'] as $item) {
            $driver->mergeCells($item);
        }

        // Copy width
        foreach ($schema['copy_width'] as $item) {
            $driver->copyWidth($item['from'], $item['to']);
        }

        // Data
        $driver->setValues($this->handleData($schema['data'], $driver, $sheetIndex), $autoCellFormat);
    }

    /**
     * @param array $data
     * @param \AnourValar\Office\Drivers\SheetsInterface $driver
     * @param int $sheetIndex
     * @return array
     */
    protected function handleData(array $data, SheetsInterface $driver, int $sheetIndex): array
    {
        foreach ($data as $row => &$columns) {
            foreach ($columns as $column => &$value) {
                if ($value instanceof \Closure) {
                    // Private Closure
                    $value = $value($driver, $column, $row);

                    if (is_null($value)) {
                        unset($data[$row][$column]);
                    }
                } elseif ($this->hookValue) {
                    // Hook: value
                    $value = ($this->hookValue)($driver, $column, $row, $value, $sheetIndex);

                    if (is_null($value)) {
                        unset($data[$row][$column]);
                    }
                }
            }
            unset($value);
        }
        unset($columns);

        return $data;
    }
}
