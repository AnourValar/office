<?php

namespace AnourValar\Office\Facades;

/**
 * Usage example:
 *
 * if (! in_array($format, [\AnourValar\Office\Format::Xlsx, \AnourValar\Office\Format::Csv])) {
 *     throw new \App\Exceptions\ValidationException('Format is not supported.');
 * }
 *
 * $generatorData = $this->buildBy($grid->query()->acl(), array_replace($this->profile, $this->profileExport)); // out of context
 * $request = $this->getBuildRequest()->get();
 *
 * return response()->streamDownload(
 *     function () use ($generatorData, $grid, $exportService, $format, $request) {
 *         echo $exportService->grid($generatorData, $grid, $format, $request);
 *     },
 *     $grid->fileName($format->fileExtension(), $request),
 *     ['Access-Control-Expose-Headers' => 'Content-Disposition']
 * );
 */

interface ExportGridQueryInterface extends ExportGridInterface
{
    /**
     * Laravel's Query builder (base query)
     *
     * @return \Illuminate\Database\Eloquent\Builder
     */
    public function query(): \Illuminate\Database\Eloquent\Builder;
}
