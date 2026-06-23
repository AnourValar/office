<?php

namespace AnourValar\Office\Tests;

use AnourValar\Office\Drivers\SheetsInterface;
use AnourValar\Office\Format;
use AnourValar\Office\Generated;
use AnourValar\Office\SheetsService;

class SheetsServiceTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Markers are replaced with data and pushed to the driver; static rows are dropped.
     *
     * @return void
     */
    public function test_generate_replaces_scalar_markers()
    {
        $driver = $this->getDriver([1 => ['A' => '[greeting]'], 2 => ['A' => 'Static text']]);

        $generated = (new SheetsService($driver))->generate('template.xlsx', ['greeting' => 'Hello']);

        $this->assertInstanceOf(Generated::class, $generated);
        $this->assertSame($driver, $generated->driver);
        $this->assertSame([1 => ['A' => 'Hello']], $driver->setValuesData);
    }

    /**
     * A closure used as a data value is invoked with (driver, column, row).
     *
     * @return void
     */
    public function test_generate_resolves_closure_values()
    {
        $driver = $this->getDriver([1 => ['A' => '[name]']]);
        $args = null;

        (new SheetsService($driver))->generate('template.xlsx', [
            'name' => function ($d, $column, $row) use (&$args, $driver) {
                $args = [$d === $driver, $column, $row];

                return "{$column}{$row}!";
            },
        ]);

        $this->assertSame([1 => ['A' => 'A1!']], $driver->setValuesData);
        $this->assertSame([true, 'A', 1], $args);
    }

    /**
     * A dynamic (list) marker grows the template: rows are added and styles copied.
     *
     * @return void
     */
    public function test_generate_expands_list_and_copies_style()
    {
        $driver = $this->getDriver([1 => ['A' => '[products.name]', 'B' => '[products.price]']]);

        (new SheetsService($driver))->generate('template.xlsx', [
            'products' => [
                ['name' => 'P1', 'price' => 10],
                ['name' => 'P2', 'price' => 20],
            ],
        ]);

        $this->assertSame(
            [
                1 => ['A' => 'P1', 'B' => 10],
                2 => ['A' => 'P2', 'B' => 20],
            ],
            $driver->setValuesData
        );
        $this->assertSame(['2:1'], $driver->calls['addRow']);
        $this->assertSame(['A1=>A2', 'B1=>B2'], $driver->calls['copyStyle']);
        $this->assertSame(['A1=>A2', 'B1=>B2'], $driver->calls['copyCellFormat']);
    }

    /**
     * A conditional marker ([=marker] with an empty value) deletes the row from the driver.
     *
     * @return void
     */
    public function test_generate_deletes_conditional_row()
    {
        $driver = $this->getDriver([
            1 => ['A' => '[=show] [title]'],
            2 => ['A' => 'Footer [other]'],
        ]);

        (new SheetsService($driver))->generate('template.xlsx', [
            'show' => false,
            'title' => 'T',
            'other' => 'O',
        ]);

        $this->assertSame(['1:1'], $driver->calls['deleteRow']);
        $this->assertSame([1 => ['A' => 'Footer O']], $driver->setValuesData);
    }

    /**
     * hookValue transforms each value; returning null removes the cell.
     *
     * @return void
     */
    public function test_hook_value_transforms_and_removes()
    {
        $driver = $this->getDriver([1 => ['A' => '[foo]', 'B' => '[bar]']]);

        (new SheetsService($driver))
            ->hookValue(function ($driver, string $column, int $row, $value, int $sheetIndex) {
                return $column === 'B' ? null : strtoupper($value);
            })
            ->generate('template.xlsx', ['foo' => 'hello', 'bar' => 'world']);

        $this->assertSame([1 => ['A' => 'HELLO']], $driver->setValuesData);
    }

    /**
     * autoCellFormat=true delegates cell formatting to the driver (no copyCellFormat).
     *
     * @return void
     */
    public function test_auto_cell_format_skips_copy_cell_format()
    {
        $driver = $this->getDriver([1 => ['A' => '[products.name]']]);

        (new SheetsService($driver))->generate(
            'template.xlsx',
            ['products' => [['name' => 'P1'], ['name' => 'P2']]],
            true
        );

        $this->assertTrue($driver->setValuesAuto);
        $this->assertArrayNotHasKey('copyCellFormat', $driver->calls);
        $this->assertSame(['A1=>A2'], $driver->calls['copyStyle']);
    }

    /**
     * @return void
     */
    public function test_iterates_every_sheet()
    {
        $driver = $this->getDriver([1 => ['A' => '[greeting]']], [], 3);

        (new SheetsService($driver))->generate('template.xlsx', ['greeting' => 'Hi']);

        $this->assertSame([0, 1, 2], $driver->sheetIndexes);
    }

    /**
     * hookLoad / hookBefore / hookAfter run in order; hookBefore can mutate data.
     *
     * @return void
     */
    public function test_hooks_order_and_before_mutation()
    {
        $driver = $this->getDriver([1 => ['A' => '[greeting]']]);
        $order = [];

        (new SheetsService($driver))
            ->hookLoad(function ($driver, $file, Format $format) use (&$order) {
                $order[] = "load:{$file}:{$format->value}";

                return $driver->load($file, $format);
            })
            ->hookBefore(function ($driver, array &$data) use (&$order) {
                $order[] = 'before';
                $data['greeting'] = 'Mutated';
            })
            ->hookAfter(function ($driver) use (&$order) {
                $order[] = 'after';
            })
            ->generate('template.xlsx', ['greeting' => 'orig']);

        $this->assertSame(['load:template.xlsx:xlsx', 'before', 'after'], $order);
        $this->assertSame([1 => ['A' => 'Mutated']], $driver->setValuesData);
    }

    /**
     * hookLoad may return a Generated; its driver is unwrapped and used.
     *
     * @return void
     */
    public function test_hook_load_unwraps_generated()
    {
        $real = $this->getDriver([1 => ['A' => '[greeting]']]);

        $generated = (new SheetsService($this->getDriver([])))
            ->hookLoad(fn ($driver, $file, $format) => new Generated($real))
            ->generate('template.xlsx', ['greeting' => 'Y']);

        $this->assertSame($real, $generated->driver);
        $this->assertSame([1 => ['A' => 'Y']], $real->setValuesData);
    }

    /**
     * @return void
     */
    public function test_canonizes_to_array_objects()
    {
        $driver = $this->getDriver([1 => ['A' => '[greeting]']]);

        $data = new class ()
        {
            public function toArray(): array
            {
                return ['greeting' => 'FromObject'];
            }
        };

        (new SheetsService($driver))->generate('template.xlsx', $data);

        $this->assertSame([1 => ['A' => 'FromObject']], $driver->setValuesData);
    }

    /**
     * @return void
     */
    public function test_format_detection()
    {
        $driver = $this->getDriver([1 => ['A' => '[greeting]']]);
        $format = null;

        $service = (new SheetsService($driver))
            ->hookLoad(function ($driver, $file, Format $f) use (&$format) {
                $format = $f;

                return $driver->load($file, $f);
            });

        $service->generate('file.ODS', ['greeting' => 'Y']);
        $this->assertSame(Format::Ods, $format);

        // unknown extension falls back to Xlsx
        $service->generate('noextension', ['greeting' => 'Y']);
        $this->assertSame(Format::Xlsx, $format);
    }

    /**
     * @param array $values
     * @param array $merge
     * @param int $sheets
     * @return object
     */
    protected function getDriver(array $values, array $merge = [], int $sheets = 1): object
    {
        return new class ($values, $merge, $sheets) implements SheetsInterface
        {
            public array $calls = [];
            public ?array $setValuesData = null;
            public ?bool $setValuesAuto = null;
            public array $sheetIndexes = [];

            public function __construct(public array $values, public array $merge, public int $sheets)
            {
                //
            }

            public function load(string $file, Format $format): SheetsInterface
            {
                return $this;
            }

            public function save(string $file, Format $format): void
            {
                //
            }

            public function setSheet(int $index): self
            {
                $this->sheetIndexes[] = $index;

                return $this;
            }

            public function getSheetCount(): int
            {
                return $this->sheets;
            }

            public function getValues(?string $ceilRange): array
            {
                return $this->values;
            }

            public function getMergeCells(): array
            {
                return $this->merge;
            }

            public function setValues(array $data, bool $autoCellFormat = true): self
            {
                $this->setValuesData = $data;
                $this->setValuesAuto = $autoCellFormat;

                return $this;
            }

            public function mergeCells(string $ceilRange): self
            {
                $this->calls['mergeCells'][] = $ceilRange;

                return $this;
            }

            public function copyStyle(string $cellFrom, string $rangeTo): self
            {
                $this->calls['copyStyle'][] = "{$cellFrom}=>{$rangeTo}";

                return $this;
            }

            public function copyCellFormat(string $cellFrom, string $rangeTo): self
            {
                $this->calls['copyCellFormat'][] = "{$cellFrom}=>{$rangeTo}";

                return $this;
            }

            public function addRow(int $rowBefore, int $qty = 1): self
            {
                $this->calls['addRow'][] = "{$rowBefore}:{$qty}";

                return $this;
            }

            public function deleteRow(int $row, int $qty = 1): self
            {
                $this->calls['deleteRow'][] = "{$row}:{$qty}";

                return $this;
            }

            public function copyWidth(string $columnFrom, string $columnTo): self
            {
                $this->calls['copyWidth'][] = "{$columnFrom}=>{$columnTo}";

                return $this;
            }
        };
    }
}
