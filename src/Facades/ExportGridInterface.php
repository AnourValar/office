<?php

namespace AnourValar\Office\Facades;

use AnourValar\Office\Drivers\GridInterface;

interface ExportGridInterface
{
    /**
     * Laravel's Query builder (base query)
     *
     * @return \Illuminate\Database\Eloquent\Builder
     */
    public function query(): \Illuminate\Database\Eloquent\Builder;

    /**
     * Columns structure
     *
     * @return array
     */
    public function columns(): array;

    /**
     * Sheet title
     *
     * @return string
     */
    public function sheetTitle(): string;

    /**
     * Row iteration
     *
     * @param \Illuminate\Database\Eloquent\Model $model
     * @param \AnourValar\Office\Drivers\GridInterface $driver
     * @param int $rowNumber
     * @return array
     */
    public function item(\Illuminate\Database\Eloquent\Model $model, GridInterface $driver, int $rowNumber): array;

    /**
     * Filename
     *
     * @param string $ext
     * @return string
     */
    public function fileName(string $ext): string;
}
