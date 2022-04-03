<?php

namespace AnourValar\Office\Drivers;

interface TemplateInterface extends SaveInterface, LoadInterface
{
    /**
     * Set values
     *
     * @param array $data
     * @return self
     */
    public function setValues(array $data): self;

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
     * Apple cell`s style to another
     *
     * @param string $cellFrom
     * @param string $rangeTo
     * @return self
     */
    public function copyStyle(string $cellFrom, string $rangeTo): self;

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
     * Set fixed width for a column
     *
     * @param string $column
     * @param mixed $width
     * @return self
     */
    public function setWidth(string $column, mixed $width = null): self;
}
