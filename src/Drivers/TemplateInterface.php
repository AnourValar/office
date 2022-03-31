<?php

namespace AnourValar\Office\Drivers;

interface TemplateInterface
{
    /**
     * Load a XLSX template
     *
     * @param string $file
     * @return self
     */
    public function loadXlsx(string $file): self;

    /**
     * Load a Ods template
     *
     * @param string $file
     * @return self
     */
    public function loadOds(string $file): self;

    /**
     * Save as XLSX
     *
     * @param string $file
     * @return void
     */
    public function saveXlsx(string $file): void;

    /**
     * Save as PDF
     *
     * @param string $file
     * @return void
     */
    public function savePdf(string $file): void;

    /**
     * Save as HTML
     *
     * @param string $file
     * @return void
     */
    public function saveHtml(string $file): void;

    /**
     * Save as ODS
     *
     * @param string $file
     * @return void
     */
    public function saveOds(string $file): void;

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
