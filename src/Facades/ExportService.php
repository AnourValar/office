<?php

namespace AnourValar\Office\Facades;

use AnourValar\Office\Drivers\GridInterface;
use AnourValar\Office\Format;

class ExportService
{
    /**
     * Extra options
     *
     * @var string
     */
    public const PERCENTAGE = 'percentage';
    public const DOUBLE_10 = 'double_10';

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
        $extras = [];

        return (new \AnourValar\Office\GridService($this->getDriver($format)))
            ->hookHeader(function (GridInterface $driver, mixed $header, string|int $key, string $column, int $rowNumber) use (&$extras) {
                if (isset($header['width'])) {
                    $driver->setWidth($column, $header['width']);
                }

                if (isset($header['height'])) {
                    $driver->setHeight($rowNumber, $header['height']);
                }

                $extras = array_merge($extras, $this->handleExtras($driver, $column, $header));

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
            ) use ($grid, &$extras) {
                $driver->setSheetTitle($grid->sheetTitle());

                foreach ($extras as $extra) {
                    $extra();
                }
            })
            ->generate($grid->columns(), $dataGenerator)
            ->save($format);
    }

    /**
     * @param \AnourValar\Office\Drivers\GridInterface $driver
     * @param string $column
     * @param array $header
     * @return array
     */
    protected function handleExtras(GridInterface $driver, string $column, array $header): array
    {
        $extras = [];

        // percentage
        if (! empty($header[self::PERCENTAGE])) {
            if ($driver instanceof \AnourValar\Office\Drivers\ZipDriver) {
                $driver->setStyle($column, 'percentage');
            } elseif ($driver instanceof \AnourValar\Office\Drivers\PhpSpreadsheetDriver) {
                $extras[] = fn () => $driver->setCellFormat($column, \AnourValar\Office\Drivers\PhpSpreadsheetDriver::FORMAT_PERCENTAGE);
            } else {
                throw new \RuntimeException('The driver does not support the "percentage" feature.');
            }
        }

        // double_10
        if (! empty($header[self::DOUBLE_10])) {
            if ($driver instanceof \AnourValar\Office\Drivers\ZipDriver) {
                $driver->setStyle($column, 'double_10');
            } elseif ($driver instanceof \AnourValar\Office\Drivers\PhpSpreadsheetDriver) {
                $extras[] = fn () => $driver->setCellFormat($column, \AnourValar\Office\Drivers\PhpSpreadsheetDriver::FORMAT_DOUBLE_10);
            } else {
                throw new \RuntimeException('The driver does not support the "double_10" feature.');
            }
        }

        return $extras;
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
