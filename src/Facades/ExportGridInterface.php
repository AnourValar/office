<?php

namespace AnourValar\Office\Facades;

use AnourValar\Office\Drivers\GridInterface;

/**
 * Usage example:
 *
 * if (! in_array($format, [\AnourValar\Office\Format::Xlsx, \AnourValar\Office\Format::Csv])) {
 *     throw new \App\Exceptions\ValidationException('Format is not supported.');
 * }
 *
 * $generatorData = $this->buildBy($myGrid->query()->acl(), array_replace($this->profile, $this->profileExport)); // outside of the stream
 *
 * return response()->streamDownload(
 *     function () use ($generatorData, $myGrid, $exportService, $format) {
 *         echo $exportService->grid($generatorData, $myGrid, $format);
 *     },
 *     $myGrid->fileName($format->fileExtension()),
 *     ['Access-Control-Expose-Headers' => 'Content-Disposition']
 * );
 */

interface ExportGridInterface
{
    /**
     * Laravel's Query builder (base query)
     *
     * @return \Illuminate\Database\Eloquent\Builder
     */
    public function query(): \Illuminate\Database\Eloquent\Builder;

    /**
     * Sheet title
     *
     * @return string
     */
    public function sheetTitle(): string;

    /**
     * Columns structure
     *
     * @return array
     */
    public function columns(): array;

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
