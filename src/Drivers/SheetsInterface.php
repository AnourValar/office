<?php

namespace AnourValar\Office\Drivers;

interface SheetsInterface extends SaveInterface, LoadInterface, MultiSheetInterface
{
    /**
     * Set values
     *
     * @param array $data
     * @param bool $autoCellFormat
     * @return self
     */
    public function setValues(array $data, bool $autoCellFormat = true): self;

    /**
     * Get values (range)
     *
     * @param string|null $ceilRange
     * @return array
     */
    public function getValues(?string $ceilRange): array;

    /**
     * Get merge cells (whole sheet)
     *
     * @return array
     */
    public function getMergeCells(): array;

    /**
     * Merge cells
     *
     * @param string $ceilRange
     * @return self
     */
    public function mergeCells(string $ceilRange): self;

    /**
     * Apply cell`s style
     *
     * @param string $cellFrom
     * @param string $rangeTo
     * @return self
     */
    public function copyStyle(string $cellFrom, string $rangeTo): self;

    /**
     * Copy cell`s format
     *
     * @param string $cellFrom
     * @param string $rangeTo
     * @return self
     */
    public function copyCellFormat(string $cellFrom, string $rangeTo): self;

    /**
     * Add a row
     *
     * @param int $rowBefore
     * @param int $qty
     * @return self
     */
    public function addRow(int $rowBefore, int $qty = 1): self;

    /**
     * Delete a row
     *
     * @param int $row
     * @param int $qty
     * @return self
     */
    public function deleteRow(int $row, int $qty = 1): self;

    /**
     * Copy column's width
     *
     * @param string $columnFrom
     * @param string $columnTo
     * @return self
     */
    public function copyWidth(string $columnFrom, string $columnTo): self;
}
