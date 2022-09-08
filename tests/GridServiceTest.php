<?php

namespace AnourValar\Office\Tests;

class GridServiceTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @return void
     */
    public function test_generate_statistic_with_headers()
    {
        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:A1', $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame('A1:A1', $totalRange);
                $this->assertSame(['A'], $columns);
            })
            ->generate(
                ['foo'],
                [ ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:A1', $headersRange);
                $this->assertSame('A2:A2', $dataRange);
                $this->assertSame('A1:A2', $totalRange);
                $this->assertSame(['A'], $columns);
            })
            ->generate(
                ['foo'],
                [ ['111'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:A1', $headersRange);
                $this->assertSame('A2:A3', $dataRange);
                $this->assertSame('A1:A3', $totalRange);
                $this->assertSame(['A'], $columns);
            })
            ->generate(
                ['foo'],
                [ ['foo-1'], ['foo-2'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:A1', $headersRange);
                $this->assertSame('A2:A4', $dataRange);
                $this->assertSame('A1:A4', $totalRange);
                $this->assertSame(['A'], $columns);
            })
            ->generate(
                ['foo'],
                [ ['foo-1'], ['foo-2'], ['foo-3'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:B1', $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame('A1:B1', $totalRange);
                $this->assertSame(['A', 'B'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:B1', $headersRange);
                $this->assertSame('A2:B2', $dataRange);
                $this->assertSame('A1:B2', $totalRange);
                $this->assertSame(['A', 'B'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ['foo-1', 'bar-1'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:B1', $headersRange);
                $this->assertSame('A2:B3', $dataRange);
                $this->assertSame('A1:B3', $totalRange);
                $this->assertSame(['A', 'B'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:B1', $headersRange);
                $this->assertSame('A2:B4', $dataRange);
                $this->assertSame('A1:B4', $totalRange);
                $this->assertSame(['A', 'B'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'], ['foo-3', 'bar-3'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:C1', $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame('A1:C1', $totalRange);
                $this->assertSame(['A', 'B', 'C'], $columns);
            })
            ->generate(
                ['foo', 'bar', 'baz'],
                [  ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:C1', $headersRange);
                $this->assertSame('A2:C2', $dataRange);
                $this->assertSame('A1:C2', $totalRange);
                $this->assertSame(['A', 'B', 'C'], $columns);
            })
            ->generate(
                ['foo', 'bar', 'baz'],
                [ ['foo-1', 'bar-1', 'baz-1']  ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:C1', $headersRange);
                $this->assertSame('A2:C3', $dataRange);
                $this->assertSame('A1:C3', $totalRange);
                $this->assertSame(['A', 'B', 'C'], $columns);
            })
            ->generate(
                ['foo', 'bar', 'baz'],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2']  ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('A1:C1', $headersRange);
                $this->assertSame('A2:C4', $dataRange);
                $this->assertSame('A1:C4', $totalRange);
                $this->assertSame(['A', 'B', 'C'], $columns);
            })
            ->generate(
                ['foo', 'bar', 'baz'],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2'], ['foo-3', 'bar-3', 'baz-3']  ],
            );
    }

    /**
     * @return void
     */
    public function test_generate_statistic_with_headers_with_shift()
    {
        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:C5', $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame('C5:C5', $totalRange);
                $this->assertSame(['one' => 'C'], $columns);
            })
            ->generate(
                ['one' => 'foo'],
                [ ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:C5', $headersRange);
                $this->assertSame('C6:C6', $dataRange);
                $this->assertSame('C5:C6', $totalRange);
                $this->assertSame(['C'], $columns);
            })
            ->generate(
                ['foo'],
                [ ['111'] ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:C5', $headersRange);
                $this->assertSame('C6:C7', $dataRange);
                $this->assertSame('C5:C7', $totalRange);
                $this->assertSame(['C'], $columns);
            })
            ->generate(
                ['foo'],
                [ ['foo-1'], ['foo-2'] ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:C5', $headersRange);
                $this->assertSame('C6:C8', $dataRange);
                $this->assertSame('C5:C8', $totalRange);
                $this->assertSame(['C'], $columns);
            })
            ->generate(
                ['foo'],
                [ ['foo-1'], ['foo-2'], ['foo-3'] ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:D5', $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame('C5:D5', $totalRange);
                $this->assertSame(['C', 'D'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:D5', $headersRange);
                $this->assertSame('C6:D6', $dataRange);
                $this->assertSame('C5:D6', $totalRange);
                $this->assertSame(['C', 'D'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ['foo-1', 'bar-1'] ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:D5', $headersRange);
                $this->assertSame('C6:D7', $dataRange);
                $this->assertSame('C5:D7', $totalRange);
                $this->assertSame(['C', 'D'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'] ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:D5', $headersRange);
                $this->assertSame('C6:D8', $dataRange);
                $this->assertSame('C5:D8', $totalRange);
                $this->assertSame(['C', 'D'], $columns);
            })
            ->generate(
                ['foo', 'bar'],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'], ['foo-3', 'bar-3'] ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:E5', $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame('C5:E5', $totalRange);
                $this->assertSame(['C', 'D', 'E'], $columns);
            })
            ->generate(
                ['foo', 'bar', 'baz'],
                [  ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:E5', $headersRange);
                $this->assertSame('C6:E6', $dataRange);
                $this->assertSame('C5:E6', $totalRange);
                $this->assertSame(['C', 'D', 'E'], $columns);
            })
            ->generate(
                ['foo', 'bar', 'baz'],
                [ ['foo-1', 'bar-1', 'baz-1']  ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:E5', $headersRange);
                $this->assertSame('C6:E7', $dataRange);
                $this->assertSame('C5:E7', $totalRange);
                $this->assertSame(['C', 'D', 'E'], $columns);
            })
            ->generate(
                ['foo', 'bar', 'baz'],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2']  ],
                'C5'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame('C5:E5', $headersRange);
                $this->assertSame('C6:E8', $dataRange);
                $this->assertSame('C5:E8', $totalRange);
                $this->assertSame(['one' => 'C', 'two' => 'D', 'three' => 'E'], $columns);
            })
            ->generate(
                ['one' => 'foo', 'two' => 'bar', 'three' => 'baz'],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2'], ['foo-3', 'bar-3', 'baz-3']  ],
                'C5'
            );
    }

    /**
     * @return void
     */
    public function test_generate_statistic_without_headers()
    {
        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame(null, $totalRange);
                $this->assertSame([], $columns);
            })
            ->generate(
                [],
                [ ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:A1', $dataRange);
                $this->assertSame('A1:A1', $totalRange);
                $this->assertSame(['A'], $columns);
            })
            ->generate(
                [],
                [ ['111'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:A2', $dataRange);
                $this->assertSame('A1:A2', $totalRange);
                $this->assertSame(['A'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1'], ['foo-2'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:A3', $dataRange);
                $this->assertSame('A1:A3', $totalRange);
                $this->assertSame(['A'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1'], ['foo-2'], ['foo-3'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:B1', $dataRange);
                $this->assertSame('A1:B1', $totalRange);
                $this->assertSame(['A', 'B'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:B2', $dataRange);
                $this->assertSame('A1:B2', $totalRange);
                $this->assertSame(['A', 'B'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:B3', $dataRange);
                $this->assertSame('A1:B3', $totalRange);
                $this->assertSame(['A', 'B'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'], ['foo-3', 'bar-3'] ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:C1', $dataRange);
                $this->assertSame('A1:C1', $totalRange);
                $this->assertSame(['A', 'B', 'C'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1', 'baz-1']  ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:C2', $dataRange);
                $this->assertSame('A1:C2', $totalRange);
                $this->assertSame(['A', 'B', 'C'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2']  ],
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('A1:C3', $dataRange);
                $this->assertSame('A1:C3', $totalRange);
                $this->assertSame(['A', 'B', 'C'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2'], ['foo-3', 'bar-3', 'baz-3']  ],
            );
    }

    /**
     * @return void
     */
    public function test_generate_statistic_without_headers_with_shift()
    {
        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame(null, $dataRange);
                $this->assertSame(null, $totalRange);
                $this->assertSame([], $columns);
            })
            ->generate(
                [],
                [ ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:D4', $dataRange);
                $this->assertSame('D4:D4', $totalRange);
                $this->assertSame(['D'], $columns);
            })
            ->generate(
                [],
                [ ['111'] ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:D5', $dataRange);
                $this->assertSame('D4:D5', $totalRange);
                $this->assertSame(['D'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1'], ['foo-2'] ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:D6', $dataRange);
                $this->assertSame('D4:D6', $totalRange);
                $this->assertSame(['D'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1'], ['foo-2'], ['foo-3'] ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:E4', $dataRange);
                $this->assertSame('D4:E4', $totalRange);
                $this->assertSame(['D', 'E'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1'] ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:E5', $dataRange);
                $this->assertSame('D4:E5', $totalRange);
                $this->assertSame(['D', 'E'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'] ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:E6', $dataRange);
                $this->assertSame('D4:E6', $totalRange);
                $this->assertSame(['D', 'E'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1'], ['foo-2', 'bar-2'], ['foo-3', 'bar-3'] ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:F4', $dataRange);
                $this->assertSame('D4:F4', $totalRange);
                $this->assertSame(['D', 'E', 'F'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1', 'baz-1']  ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:F5', $dataRange);
                $this->assertSame('D4:F5', $totalRange);
                $this->assertSame(['D', 'E', 'F'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2']  ],
                'D4'
            );

        (new \AnourValar\Office\GridService($this->getDriver()))
            ->hookAfter(function ($driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns) {
                $this->assertSame(null, $headersRange);
                $this->assertSame('D4:F6', $dataRange);
                $this->assertSame('D4:F6', $totalRange);
                $this->assertSame(['D', 'E', 'F'], $columns);
            })
            ->generate(
                [],
                [ ['foo-1', 'bar-1', 'baz-1'], ['foo-2', 'bar-2', 'baz-2'], ['foo-3', 'bar-3', 'baz-3']  ],
                'D4'
            );
    }

    /**
     * @return \AnourValar\Office\Drivers\GridInterface
     */
    protected function getDriver(): \AnourValar\Office\Drivers\GridInterface
    {
        return new class implements \AnourValar\Office\Drivers\GridInterface
        {
            public function create(): self
            {
                return $this;
            }


            public function setGrid(iterable $data): self
            {
                foreach ($data as $item) {

                }

                return $this;
            }

            public function save(string $file, \AnourValar\Office\Format $format): void
            {

            }
        };
    }
}
