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
 * $request = $this->getBuildRequest()->get();
 *
 * return response()->streamDownload(
 *     function () use ($generatorData, $myGrid, $exportService, $format, $request) {
 *         echo $exportService->grid($generatorData, $myGrid, $format, $request);
 *     },
 *     $myGrid->fileName($format->fileExtension(), $request),
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
     * @param \Illuminate\Database\Eloquent\Model $model
     * @param \AnourValar\Office\Drivers\GridInterface $driver
     * @param int $rowNumber
     * @param array $request
     * @return array
     */
    public function item(\Illuminate\Database\Eloquent\Model $model, GridInterface $driver, int $rowNumber, array $request): array;

    /**
     * Filename
     *
     * @param string $ext
     * @param array $request
     * @return string
     */
    public function fileName(string $ext, array $request): string;
}
