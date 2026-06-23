<?php

namespace AnourValar\Office\Tests;

use AnourValar\Office\Drivers\GridInterface;
use AnourValar\Office\Format;
use AnourValar\Office\Generated;
use AnourValar\Office\GridService;

class GridServiceHooksTest extends \PHPUnit\Framework\TestCase
{
    /**
     * The generator feeds headers followed by data rows to the driver.
     *
     * @return void
     */
    public function test_generate_emits_headers_and_data()
    {
        $driver = $this->getDriver();

        (new GridService($driver))->generate(
            ['Name', 'Sales'],
            [ ['William', 3000], ['James', 4000] ],
        );

        $this->assertSame(
            [
                ['Name', 'Sales'],
                ['William', 3000],
                ['James', 4000],
            ],
            $driver->grid
        );
    }

    /**
     * @return void
     */
    public function test_generate_without_headers_emits_only_data()
    {
        $driver = $this->getDriver();

        (new GridService($driver))->generate(
            [],
            [ ['William', 3000], ['James', 4000] ],
        );

        $this->assertSame(
            [
                ['William', 3000],
                ['James', 4000],
            ],
            $driver->grid
        );
    }

    /**
     * @return void
     */
    public function test_generate_accepts_closure_data()
    {
        $driver = $this->getDriver();

        (new GridService($driver))->generate(
            ['Name'],
            function () {
                yield ['William'];
                yield ['James'];
            },
        );

        $this->assertSame([['Name'], ['William'], ['James']], $driver->grid);
    }

    /**
     * @return void
     */
    public function test_left_top_corner_indents_with_empty_cells_and_rows()
    {
        $driver = $this->getDriver();

        (new GridService($driver))->generate(
            ['Name'],
            [ ['William'] ],
            'C2'
        );

        // one leading empty row (row 1), then 'C' column shifted by two empty cells
        $this->assertSame(
            [
                [],
                ['', '', 'Name'],
                ['', '', 'William'],
            ],
            $driver->grid
        );
    }

    /**
     * @return void
     */
    public function test_hook_header_transforms_and_receives_position()
    {
        $driver = $this->getDriver();
        $seen = [];

        (new GridService($driver))
            ->hookHeader(function (GridInterface $driver, $header, $key, string $column, int $rowNumber) use (&$seen) {
                $seen[] = [$header['title'], $key, $column, $rowNumber];

                return $header['title'];
            })
            ->generate(
                [ ['title' => 'Name'], ['title' => 'Sales'] ],
                [ ['William', 3000] ],
            );

        $this->assertSame([['Name'], ['Sales']], array_map(fn ($x) => [$x[0]], $seen));
        $this->assertSame(
            [
                ['Name', 0, 'A', 1],
                ['Sales', 1, 'B', 1],
            ],
            $seen
        );
        $this->assertSame(['Name', 'Sales'], $driver->grid[0]);
    }

    /**
     * @return void
     */
    public function test_hook_row_transforms_and_receives_position()
    {
        $driver = $this->getDriver();
        $rows = [];

        (new GridService($driver))
            ->hookRow(function (GridInterface $driver, $row, $key, int $rowNumber) use (&$rows) {
                $rows[] = [$key, $rowNumber];

                return [$row['name'], $row['sales']];
            })
            ->generate(
                ['Name', 'Sales'],
                [ ['name' => 'William', 'sales' => 3000], ['name' => 'James', 'sales' => 4000] ],
            );

        $this->assertSame([[0, 2], [1, 3]], $rows);
        $this->assertSame(
            [
                ['Name', 'Sales'],
                ['William', 3000],
                ['James', 4000],
            ],
            $driver->grid
        );
    }

    /**
     * Returning null from hookRow skips the row entirely (and the row number is not consumed).
     *
     * @return void
     */
    public function test_hook_row_null_skips_row()
    {
        $driver = $this->getDriver();
        $seenRowNumbers = [];

        (new GridService($driver))
            ->hookRow(function (GridInterface $driver, $row, $key, int $rowNumber) use (&$seenRowNumbers) {
                $seenRowNumbers[] = $rowNumber;

                return $row[0] === 'skip' ? null : $row;
            })
            ->hookAfter(function ($driver, $headersRange, $dataRange, $totalRange) {
                $this->assertSame('A2:A3', $dataRange);
            })
            ->generate(
                ['Name'],
                [ ['William'], ['skip'], ['James'] ],
            );

        $this->assertSame([['Name'], ['William'], ['James']], $driver->grid);
        // William consumes row 2; the skipped row does not advance the counter, so James also gets 3
        $this->assertSame([2, 3, 3], $seenRowNumbers);
    }

    /**
     * @return void
     */
    public function test_hook_before_receives_arguments()
    {
        $driver = $this->getDriver();
        $captured = null;

        (new GridService($driver))
            ->hookBefore(function ($driver, array &$headers, $data, string $leftTopCorner) use (&$captured) {
                $captured = [$headers, $leftTopCorner];
            })
            ->generate(['Name'], [ ['William'] ], 'B2');

        $this->assertSame([['Name'], 'B2'], $captured);
    }

    /**
     * hookLoad replaces the default driver creation.
     *
     * @return void
     */
    public function test_hook_load_replaces_driver()
    {
        $original = $this->getDriver();
        $replacement = $this->getDriver();

        $generated = (new GridService($original))
            ->hookLoad(function ($driver) use ($original, $replacement) {
                $this->assertSame($original, $driver);

                return $replacement->create();
            })
            ->generate(['Name'], [ ['William'] ]);

        $this->assertSame($replacement, $generated->driver);
        $this->assertSame([['Name'], ['William']], $replacement->grid);
    }

    /**
     * hookLoad may return a Generated, whose driver is unwrapped and used.
     *
     * @return void
     */
    public function test_hook_load_unwraps_generated()
    {
        $replacement = $this->getDriver();

        $generated = (new GridService($this->getDriver()))
            ->hookLoad(fn ($driver) => new Generated($replacement->create()))
            ->generate(['Name'], [ ['William'] ]);

        $this->assertSame($replacement, $generated->driver);
        $this->assertSame([['Name'], ['William']], $replacement->grid);
    }

    /**
     * @return object
     */
    protected function getDriver(): object
    {
        return new class () implements GridInterface
        {
            public array $grid = [];

            public function create(): self
            {
                return $this;
            }

            public function setGrid(iterable $data): self
            {
                foreach ($data as $row) {
                    $this->grid[] = $row;
                }

                return $this;
            }

            public function save(string $file, Format $format): void
            {
                //
            }
        };
    }
}
