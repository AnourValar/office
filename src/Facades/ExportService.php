<?php

namespace AnourValar\Office\Facades;

use AnourValar\Office\Drivers\GridInterface;
use AnourValar\Office\Format;

class ExportService
{
    /**
     * Generate a grid
     *
     * @param \Closure $dataGenerator
     * @param \AnourValar\Office\Facades\ExportGridInterface $grid
     * @param \AnourValar\Office\Format $format
     * @return string
     */
    public function grid(\Closure $dataGenerator, ExportGridInterface $grid, Format $format = Format::Xlsx): string
    {
        return (new \AnourValar\Office\GridService($this->getDriver($format)))
            ->hookHeader(function (GridInterface $driver, mixed $header, string|int $key, string $column, int $rowNumber) {
                if (isset($header['width'])) {
                    $driver->setWidth($column, $header['width']);
                }

                if (isset($header['height'])) {
                    $driver->setHeight($rowNumber, $header['height']);
                }

                return $header['title'];
            })
            ->hookRow(function (GridInterface $driver, mixed $row, string|int $key, int $rowNumber) use ($grid) {
                return $grid->item($row, $driver, $rowNumber);
            })
            ->hookAfter(function (
                GridInterface $driver,
                ?string $headersRange,
                ?string $dataRange,
                ?string $totalRange,
                array $columns
            ) use ($grid) {
                $driver->setSheetTitle($grid->sheetTitle());
            })
            ->generate($grid->columns(), $dataGenerator)
            ->save($format);
    }

    /**
     * @param \AnourValar\Office\Format $format
     * @return \AnourValar\Office\Drivers\GridInterface
     */
    protected function getDriver(Format $format): GridInterface
    {
        if ($format == Format::Xlsx) {
            return new \AnourValar\Office\Drivers\ZipDriver();
        }

        return new \AnourValar\Office\Drivers\PhpSpreadsheetDriver();
    }
}
