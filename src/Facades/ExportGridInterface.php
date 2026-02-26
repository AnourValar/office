<?php

namespace AnourValar\Office\Facades;

use AnourValar\Office\Drivers\GridInterface;

interface ExportGridInterface
{
    /**
     * Sheet title
     *
     * @param array $request
     * @return string
     */
    public function sheetTitle(array $request): string;

    /**
     * Columns structure
     *
     * @param array $request
     * @return array
     */
    public function columns(array $request): array;

    /**
     * Row iteration
     *
     * @param mixed $row
     * @param \AnourValar\Office\Drivers\GridInterface $driver
     * @param int $rowNumber
     * @param array $request
     * @return array
     */
    public function item($row, GridInterface $driver, int $rowNumber, array $request): array;

    /**
     * Filename
     *
     * @param string $ext
     * @param array $request
     * @return string
     */
    public function fileName(string $ext, array $request): string;
}
