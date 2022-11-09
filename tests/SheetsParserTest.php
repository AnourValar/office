<?php

namespace AnourValar\Office\Tests;

class SheetsParserTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @var \AnourValar\Office\Sheets\Parser
     */
    protected \AnourValar\Office\Sheets\Parser $service;

    /**
     * @see \PHPUnit\Framework\TestCase
     *
     * @return void
     */
    protected function setUp(): void
    {
        $this->service = new \AnourValar\Office\Sheets\Parser();
    }

    /**
     * @return void
     */
    public function test_schema_scalar()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => 'test [foo] test',
                        'E' => '[test10]',
                        'F' => '[test]',
                        'G' => '[test1]',
                        'H' => 'hello [$=foo]',
                        'I' => '[$= foo] world',
                        'J' => '[$! foo] test',
                        'K' => '[$! foo] test',
                        'L' => '[$= foo1] test',
                        'Q' => '=A1+B2+C3+D4+E5',
                        'X' => '= A1',
                        'W' => '=1',
                        'Y' => '=AB100',
                    ],
                    2 => [
                        'A' => '[bar]',
                        'B' => 'test [bar] test',
                        'C' => 'bar',
                        'D' => '[a.b] -> [a.c] -> [a.d] -> [a.e.f]',
                        'H' => '[hello]',
                        'J' => '[world]',
                        'K' => '[k] [9] [9k] [k9]',
                    ],
                    '3' => [
                        'D' => '[bar] [!bar]',
                    ],
                    '4' => [
                        'D' => 'hello world',
                    ],
                    '5' => [
                        'A' => 1,
                        'B' => 2,
                    ],
                    '6' => [
                        'Q' => '=A1+B2+C3+D4+E5',
                    ],
                ],

                'data' => [
                    'foo' => 'hello',
                    'bar' => 'world',
                    'a' => ['b' => '11', 'c' => '22', 'e' => ['f' => 'oops']],
                    'test10' => 12.5,
                    'test' => '3',
                    'test1' => 5,
                    'hello' => null,
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'hello',
                            'C' => 'test hello test',
                            'E' => 12.5,
                            'F' => '3',
                            'G' => 5,
                            'H' => 'hello',
                            'I' => 'world',
                            'J' => null,
                            'K' => null,
                            'L' => null,
                            'Q' => '=A1+B2+C3+D3+E4',
                            'Y' => '=AB99',
                        ],
                        2 => [
                            'A' => 'world',
                            'B' => 'test world test',
                            'D' => '11 -> 22 ->  -> oops',
                            'H' => null,
                            'J' => null,
                            'K' => '[k] [9] [9k]',
                        ],
                        5 => [
                            'Q' => '=A1+B2+C3+D3+E4',
                        ],
                    ],

                    'rows' => [['action' => 'delete', 'row' => 3]],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_closure()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                    ],
                    2 => [
                        'A' => 'hello [bar] world',
                    ],
                    3 => [
                        'A' => '[test] [=foo]',
                    ],
                    4 => [
                        'A' => '[test2] [=bar]',
                    ],
                    5 => [
                        'A' => '[test2] [!foo]',
                    ],
                ],

                'data' => [
                    'foo' => function () {},
                    'test' => function () {},
                    'test2' => function () { throw new \LogicException('oops'); },
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => $item['data']['foo'],
                        ],
                        2 => [
                            'A' => 'hello  world',
                        ],
                        3 => [
                            'A' => $item['data']['test'],
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 4],
                        ['action' => 'delete', 'row' => 4],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }

        $this->assertSame(
            [
                'data' => [
                    1 => [
                        'A' => 'hello',
                    ],
                ],

                'rows' => [],

                'copy_style' => [],

                'merge_cells' => [],

                'copy_width' => [],

                'copy_cell_format' => [],
            ],
            $this->service->schema([1 => ['A' => 'hello [world]']], ['world' => function () {}], [])->toArray()
        );
    }

    /**
     * @return void
     */
    public function test_schema_list_zero_empty()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo] [=test]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo] [= list]',
                    ],
                    3 => [
                        'A' => 'bar [= bar]',
                        'B' => '[bar] 111',
                    ],
                    4 => [
                        'A' => 'foo [= list.0]',
                        'B' => '[foo]',
                    ],
                    5 => [
                        'A' => 'bar',
                        'B' => '[bar] 222 [=bar]',
                        'C' => '=A3+B5+C7',
                    ],
                    6 => [
                        'A' => 'foo [= list.0.c]',
                        'B' => '[foo] [!list]',
                    ],
                    7 => [
                        'A' => 'bar',
                        'B' => '[bar] 333 [! list]',
                    ],
                ],

                'data' => [
                    'list' => [],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'bar',
                            'B' => 'test2 111',
                        ],
                        2 => [
                            'B' => 'test2 222',
                            'C' => '=A1+B2+C3',
                        ],
                        3 => [
                            'B' => 'test2 333',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 1],
                        ['action' => 'delete', 'row' => 1],
                        ['action' => 'delete', 'row' => 2],
                        ['action' => 'delete', 'row' => 3],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "#$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_zero()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo [= list.0.c]',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.0.c]',
                        'D' => '[list.0.d]',
                        'E' => '[foo] hello [list.0.c] -> [list.0.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'I' => '=A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A1*B2*C3',
                    ],
                ],

                'data' => [
                    'list' => [],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo] [=list_c.0]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo] [!list_c]',

                        'C' => '[list_c.0]',
                        'D' => '[list_d.0]',
                        'E' => '[foo] hello [list_c.0] -> [list_d.0] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar] [! list_c]',
                        'I' => '=A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A1*B2*C3',
                    ],
                ],

                'data' => [
                    'list_c' => [],
                    'list_d' => [],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [

                        1 => [
                            'B' => 'test1',
                            'C' => null,
                            'D' => null,
                            'E' => 'test1 hello  ->  world test2',
                            'G' => 'test2',
                            'I' => '=A1*B1*C2',
                            'K' => '=A1:A1',
                        ],
                        '2' => [
                            'B' => 'test2',
                            'I' => '=A1*B1*C2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 1],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "#$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_one()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.0.c]',
                        'D' => '[list.0.d]',
                        'E' => '[foo] hello [list.0.c] -> [list.0.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    4 => [
                        'A' => '[bar] [! list]',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => '11', 'd' => '12'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.c]',
                        'D' => '[list.d]',
                        'E' => '[foo] hello [list.c] -> [list.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    4 => [
                        'A' => '[bar] [!list]',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => '11', 'd' => '12'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c.0]',
                        'D' => '[list_d.0]',
                        'E' => '[foo] hello [list_c.0] -> [list_d.0] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    4 => [
                        'A' => '[bar] [! list_c]',
                    ],
                ],

                'data' => [
                    'list_c' => ['11'],
                    'list_d' => ['12'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c]',
                        'D' => '[list_d]',
                        'E' => '[foo] hello [list_c] -> [list_d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    4 => [
                        'A' => '[bar] [!list_c]',
                    ],
                ],

                'data' => [
                    'list_c' => ['11'],
                    'list_d' => ['12'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'H' => '=A1*B2*C3',
                            'K' => '=A2:A2',
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => '11',
                            'D' => '12',
                            'E' => 'test1 hello 11 -> 12 world test2',
                            'G' => 'test2',
                            'H' => '=A1*B2*C3',
                        ],
                        3 => [
                            'B' => 'test2',
                            'H' => '=A1*B2*C3',
                            'K' => '=A2:A2',
                        ],
                    ],

                    'rows' => [['action' => 'delete', 'row' => 4]],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_two()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.c]',
                        'D' => '[list.d]',
                        'E' => '[foo] hello [list.c] -> [list.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => '11', 'd' => '12'], ['c' => '21', 'd' => '22'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.c]',
                        'D' => '[list.d]',
                        'E' => '[foo] hello [list.*.c] -> [list.*.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => '11', 'd' => '12'], ['c' => '21', 'd' => '22'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c]',
                        'D' => '[list_d]',
                        'E' => '[foo] hello [list_c] -> [list_d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list_c' => ['11', '21'],
                    'list_d' => ['12', '22'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c.*]',
                        'D' => '[list_d.*]',
                        'E' => '[foo] hello [list_c.*] -> [list_d.*] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'H' => '=A1*B2*C3',
                        'J' => 'A1*B2*C3',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list_c' => ['11', '21'],
                    'list_d' => ['12', '22'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'H' => '=A1*B2*C4',
                            'K' => '=A2:A3',
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '11',
                            'D' => '12',
                            'E' => 'test1 hello 11 -> 12 world test2',
                            'F' => 'bar',
                            'G' => 'test2',
                            'H' => '=A1*B2*C4',
                            'J' => 'A1*B2*C3',
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '21',
                            'D' => '22',
                            'E' => 'test1 hello 21 -> 22 world test2',
                            'F' => 'bar',
                            'G' => 'test2',
                            'H' => '=A1*B3*C4',
                            'J' => 'A1*B2*C3',
                        ],
                        4 => [
                            'B' => 'test2',
                            'H' => '=A1*B2*C4',
                            'K' => '=A2:A3',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'J2', 'to' => 'J3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'J2', 'to' => 'J3'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_three()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.c]',
                        'D' => '[list.d]',
                        'E' => '[foo] hello [list.c] -> [list.d] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => 11, 'd' => 12.5], ['c' => '21', 'd' => '22'], ['c' => '31', 'd' => '32'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.*.c]',
                        'D' => '[list.*.d]',
                        'E' => '[foo] hello [list.*.c] -> [list.*.d] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => 11, 'd' => 12.5], ['c' => '21', 'd' => '22'], ['c' => '31', 'd' => '32'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c]',
                        'D' => '[list_d]',
                        'E' => '[foo] hello [list_c] -> [list_d] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list_c' => [11, '21', '31'],
                    'list_d' => [12.5, '22', '32'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c.*]',
                        'D' => '[list_d.*]',
                        'E' => '[foo] hello [list_c.*] -> [list_d.*] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list_c' => [11, '21', '31'],
                    'list_d' => [12.5, '22', '32'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'I' => '=A2*C2+B1+D5+E$2',
                            'K' => '=A2:A4',
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => 11,
                            'D' => 12.5,
                            'E' => 'test1 hello 11 -> 12.5 world test2',
                            'G' => 'bar',
                            'H' => 'test2',
                            'I' => '=A2*C2+B1+D5+E$2',
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '21',
                            'D' => '22',
                            'E' => 'test1 hello 21 -> 22 world test2',
                            'G' => 'bar',
                            'H' => 'test2',
                            'I' => '=A3*C3+B1+D5+E$2',
                        ],
                        4 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '31',
                            'D' => '32',
                            'E' => 'test1 hello 31 -> 32 world test2',
                            'G' => 'bar',
                            'H' => 'test2',
                            'I' => '=A4*C4+B1+D5+E$2',
                        ],
                        5 => [
                            'B' => 'test2',
                            'I' => '=A2*C2+B1+D5+E$2',
                            'K' => '=A2:A4',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'F2', 'to' => 'F4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'H2', 'to' => 'H4'],
                        ['from' => 'I2', 'to' => 'I3'],
                        ['from' => 'I2', 'to' => 'I4'],
                    ],

                    'merge_cells' => ['E3:F3', 'E4:F4'],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'H2', 'to' => 'H4'],
                        ['from' => 'I2', 'to' => 'I3'],
                        ['from' => 'I2', 'to' => 'I4'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_three_limit()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.0.c]',
                        'D' => '[list.0.d]',
                        'E' => '[foo] hello [list.0.c] -> [list.0.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => 11, 'd' => 12.5], ['c' => '21', 'd' => '22'], ['c' => '31', 'd' => '32'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c.0]',
                        'D' => '[list_d.0]',
                        'E' => '[foo] hello [list_c.0] -> [list_d.0] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'list_c' => [11, '21', '31'],
                    'list_d' => [12.5, '22', '32'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => 11,
                            'D' => 12.5,
                            'E' => 'test1 hello 11 -> 12.5 world test2',
                            'G' => 'test2',
                        ],

                        3 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_four()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.c]',
                        'D' => '[list.d]',
                        'E' => '[foo] hello [list.c] -> [list.d] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => 11, 'd' => 12.5], ['c' => '21', 'd' => '22'], ['c' => '31', 'd' => '32'], ['c' => '41', 'd' => '42'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.*.c]',
                        'D' => '[list.*.d]',
                        'E' => '[foo] hello [list.*.c] -> [list.*.d] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list' => [ ['c' => 11, 'd' => 12.5], ['c' => '21', 'd' => '22'], ['c' => '31', 'd' => '32'], ['c' => '41', 'd' => '42'] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c]',
                        'D' => '[list_d]',
                        'E' => '[foo] hello [list_c] -> [list_d] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list_c' => [11, '21', '31', '41'],
                    'list_d' => [12.5, '22', '32', '42'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c.*]',
                        'D' => '[list_d.*]',
                        'E' => '[foo] hello [list_c.*] -> [list_d.*] world [bar]',

                        'G' => 'bar',
                        'H' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'I' => '=A2*C2+B1+D3+E$2',
                        'K' => '=A2:A2',
                    ],
                ],

                'data' => [
                    'list_c' => [11, '21', '31', '41'],
                    'list_d' => [12.5, '22', '32', '42'],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['E2:F2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'I' => '=A2*C2+B1+D6+E$2',
                            'K' => '=A2:A5',
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => 11,
                            'D' => 12.5,
                            'E' => 'test1 hello 11 -> 12.5 world test2',
                            'G' => 'bar',
                            'H' => 'test2',
                            'I' => '=A2*C2+B1+D6+E$2',
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '21',
                            'D' => '22',
                            'E' => 'test1 hello 21 -> 22 world test2',
                            'G' => 'bar',
                            'H' => 'test2',
                            'I' => '=A3*C3+B1+D6+E$2',
                        ],
                        4 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '31',
                            'D' => '32',
                            'E' => 'test1 hello 31 -> 32 world test2',
                            'G' => 'bar',
                            'H' => 'test2',
                            'I' => '=A4*C4+B1+D6+E$2',
                        ],
                        5 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '41',
                            'D' => '42',
                            'E' => 'test1 hello 41 -> 42 world test2',
                            'G' => 'bar',
                            'H' => 'test2',
                            'I' => '=A5*C5+B1+D6+E$2',
                        ],
                        6 => [
                            'B' => 'test2',
                            'I' => '=A2*C2+B1+D6+E$2',
                            'K' => '=A2:A5',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'E2', 'to' => 'E5'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'F2', 'to' => 'F4'],
                        ['from' => 'F2', 'to' => 'F5'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                        ['from' => 'G2', 'to' => 'G5'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'H2', 'to' => 'H4'],
                        ['from' => 'H2', 'to' => 'H5'],
                        ['from' => 'I2', 'to' => 'I3'],
                        ['from' => 'I2', 'to' => 'I4'],
                        ['from' => 'I2', 'to' => 'I5'],
                    ],

                    'merge_cells' => ['E3:F3', 'E4:F4', 'E5:F5'],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'E2', 'to' => 'E5'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                        ['from' => 'G2', 'to' => 'G5'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'H2', 'to' => 'H4'],
                        ['from' => 'H2', 'to' => 'H5'],
                        ['from' => 'I2', 'to' => 'I3'],
                        ['from' => 'I2', 'to' => 'I4'],
                        ['from' => 'I2', 'to' => 'I5'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_zero_empty()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.c] [= test]',
                    ],
                    3 => [
                        'A' => 'bar [= test]',
                        'B' => '[foo]',
                    ],
                    4 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => []] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo [= test]',
                        'B' => '[foo]',

                        'C' => '[matrix.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[foo] [= test]',
                    ],
                    4 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [[]] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'C' => '=SUM(A2:A2)',
                        ],
                        2 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 2],
                        ['action' => 'delete', 'row' => 2],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "#$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_zero()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => []] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => []] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [[]] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [[]] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'C' => '=SUM(A2:A2)',
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => null,
                        ],
                        '3' => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "#$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_one()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.c.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.*.c.*]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.*.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.0.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [['11']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.*.*.*]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [['11']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [['11']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'C' => '=SUM(A2:A2)',
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => '11',
                        ],
                        3 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_two()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11', '12']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.*.c.*]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11', '12']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.*.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11', '12']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [['11', '12']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'C' => '=SUM(A2:A2)',
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => '11',
                            'D' => '12',
                        ],
                        3 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [
                        ['from' => 'C2', 'to' => 'D2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [['from' => 'C', 'to' => 'D']],

                    'copy_cell_format' => [
                        ['from' => 'C2', 'to' => 'D2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_two_limit()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.c.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11', '12']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.0.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [['11', '12']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => '11',
                        ],
                        3 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_three()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11', '12', '13']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.*.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11', '12', '13']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.*.c.*]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ ['c' => ['11', '12', '13']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'matrix' => [ [['11', '12', '13']] ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'C' => '=SUM(A2:A2)',
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => '11',
                            'D' => '12',
                            'E' => '13',
                        ],
                        3 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'C', 'to' => 'D'],
                        ['from' => 'C', 'to' => 'E'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_equal()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.b]',
                        'C' => '[matrix.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => 'foo: [foo] bar: [bar]',
                        'D' => 'bar: [bar] foo: [foo]',
                    ],
                ],

                'data' => [
                    'matrix' => [
                        ['b' => 'b1', 'c' => ['c1', 'd1', 'e1']],
                        ['b' => 'b2', 'c' => ['c2', 'd2', 'e2']],
                        ['b' => 'b3', 'c' => ['c3', 'd3', 'e3']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.*.b]',
                        'C' => '[matrix.*.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => 'foo: [foo] bar: [bar]',
                        'D' => 'bar: [bar] foo: [foo]',
                    ],
                ],

                'data' => [
                    'matrix' => [
                        ['b' => 'b1', 'c' => ['c1', 'd1', 'e1']],
                        ['b' => 'b2', 'c' => ['c2', 'd2', 'e2']],
                        ['b' => 'b3', 'c' => ['c3', 'd3', 'e3']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'C' => '=SUM(A2:A2)',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.*.b.*]',
                        'C' => '[matrix.*.c.*]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => 'foo: [foo] bar: [bar]',
                        'D' => 'bar: [bar] foo: [foo]',
                    ],
                ],

                'data' => [
                    'matrix' => [
                        ['b' => 'b1', 'c' => ['c1', 'd1', 'e1']],
                        ['b' => 'b2', 'c' => ['c2', 'd2', 'e2']],
                        ['b' => 'b3', 'c' => ['c3', 'd3', 'e3']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'C' => '=SUM(A2:A4)',
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'b1',
                            'C' => 'c1',
                            'D' => 'd1',
                            'E' => 'e1',
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'b2',
                            'C' => 'c2',
                            'D' => 'd2',
                            'E' => 'e2',
                        ],
                        4 => [
                            'A' => 'foo',
                            'B' => 'b3',
                            'C' => 'c3',
                            'D' => 'd3',
                            'E' => 'e3',
                        ],
                        5 => [
                            'B' => 'test2',
                            'C' => 'foo: test1 bar: test2',
                            'D' => 'bar: test2 foo: test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'C', 'to' => 'D'],
                        ['from' => 'C', 'to' => 'E'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_equal_limit()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.0.b]',
                        'C' => '[matrix.0.c.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => 'foo: [foo] bar: [bar]',
                        'D' => 'bar: [bar] foo: [foo]',
                    ],
                ],

                'data' => [
                    'matrix' => [
                        ['b' => 'b1', 'c' => ['c1', 'd1', 'e1']],
                        ['b' => 'b2', 'c' => ['c2', 'd2', 'e2']],
                        ['b' => 'b3', 'c' => ['c3', 'd3', 'e3']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                        ],
                        2 => [
                            'B' => 'b1',
                            'C' => 'c1',
                        ],
                        3 => [
                            'B' => 'test2',
                            'C' => 'foo: test1 bar: test2',
                            'D' => 'bar: test2 foo: test1',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_irr()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.0.b] [= matrix.0.b]',
                        'C' => '[matrix.0.c.0] [=matrix.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => 'foo: [foo] bar: [bar]',
                        'D' => 'bar: [bar] foo: [foo]',
                    ],
                ],

                'data' => [
                    'matrix' => [
                        ['b' => 'b1', 'c' => ['c1', 'd1', 'e1']],
                        ['b' => 'b2', 'c' => ['c2', 'd2']],
                        ['c' => ['c3', 'd3', 'e3']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.b]',
                        'C' => '[matrix.c]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => 'foo: [foo] bar: [bar]',
                        'D' => 'bar: [bar] foo: [foo]',
                    ],
                ],

                'data' => [
                    'matrix' => [
                        ['b' => 'b1', 'c' => ['c1', 'd1', 'e1']],
                        ['b' => 'b2', 'c' => ['c2', 'd2']],
                        ['c' => ['c3', 'd3', 'e3']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'b1',
                            'C' => 'c1',
                            'D' => 'd1',
                            'E' => 'e1',
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'b2',
                            'C' => 'c2',
                            'D' => 'd2',
                            'E' => null,
                        ],
                        4 => [
                            'A' => 'foo',
                            'B' => null,
                            'C' => 'c3',
                            'D' => 'd3',
                            'E' => 'e3',
                        ],
                        5 => [
                            'B' => 'test2',
                            'C' => 'foo: test1 bar: test2',
                            'D' => 'bar: test2 foo: test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'C', 'to' => 'D'],
                        ['from' => 'C', 'to' => 'E'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_combination1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.b]',
                        'C' => '[rows.c]',
                        'E' => '[columns.e]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'rows' => [
                        ['b' => 'b1', 'c' => 'c1'],
                        ['b' => 'b2', 'c' => 'c2'],
                    ],
                    'columns' => [
                        ['e' => ['E10', 'E11', '']],
                        ['e' => ['E20', 'E21', 'E22']],
                        ['e' => ['E30', 'E31', 'E32']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.*.b]',
                        'C' => '[rows.*.c]',
                        'E' => '[columns.*.e.*]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'rows' => [
                        ['b' => 'b1', 'c' => 'c1'],
                        ['b' => 'b2', 'c' => 'c2'],
                    ],
                    'columns' => [
                        ['e' => ['E10', 'E11', '']],
                        ['e' => ['E20', 'E21', 'E22']],
                        ['e' => ['E30', 'E31', 'E32']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'b1',
                            'C' => 'c1',
                            'E' => 'E10',
                            'F' => 'E11',
                            'G' => null,
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'b2',
                            'C' => 'c2',
                            'E' => 'E20',
                            'F' => 'E21',
                            'G' => 'E22',
                        ],
                        4 => [
                            'A' => 'foo',
                            'B' => null,
                            'C' => null,
                            'E' => 'E30',
                            'F' => 'E31',
                            'G' => 'E32',
                        ],
                        5 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'E2', 'to' => 'F2'],
                        ['from' => 'E2', 'to' => 'G2'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'F2', 'to' => 'F4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'E', 'to' => 'F'],
                        ['from' => 'E', 'to' => 'G'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'E2', 'to' => 'F2'],
                        ['from' => 'E2', 'to' => 'G2'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'F2', 'to' => 'F4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_combination1_limit()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.0.b]',
                        'C' => '[rows.0.c]',
                        'E' => '[columns.0.e.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                    ],
                ],

                'data' => [
                    'rows' => [
                        ['b' => 'b1', 'c' => 'c1'],
                        ['b' => 'b2', 'c' => 'c2'],
                    ],
                    'columns' => [
                        ['e' => ['E10', 'E11', '']],
                        ['e' => ['E20', 'E21', 'E22']],
                        ['e' => ['E30', 'E31', 'E32']],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                        ],
                        2 => [
                            'B' => 'b1',
                            'C' => 'c1',
                            'E' => 'E10',
                        ],
                        3 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }


    /**
     * @return void
     */
    public function test_schema_multi_combination2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'E' => '[columns.one]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.b]',
                        'C' => '[rows.c]',
                        'E' => '[columns.two]',
                    ],
                    3 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'rows' => [
                        ['b' => 'b1', 'c' => 'c1'],
                        ['b' => 'b2', 'c' => 'c2'],
                    ],
                    'columns' => [
                        [
                            'one' => ['01', '02', '03'],
                            'two' => [15000, 20000, 30000],
                        ],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['C2:D2'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'E' => '[columns.*.one]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.b]',
                        'C' => '[rows.c]',
                        'E' => '[columns.*.two.*]',
                    ],
                    3 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'rows' => [
                        ['b' => 'b1', 'c' => 'c1'],
                        ['b' => 'b2', 'c' => 'c2'],
                    ],
                    'columns' => [
                        [
                            'one' => ['01', '02', '03'],
                            'two' => [15000, 20000, 30000],
                        ],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],

                'merge_cells' => ['C2:D2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'E' => '01',
                            'F' => '02',
                            'G' => '03',
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'b1',
                            'C' => 'c1',
                            'E' => 15000,
                            'F' => 20000,
                            'G' => 30000,
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'b2',
                            'C' => 'c2',
                            'E' => null,
                            'F' => null,
                            'G' => null,
                        ],
                        4 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'E1', 'to' => 'F1'],
                        ['from' => 'E1', 'to' => 'G1'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'F2'],
                        ['from' => 'E2', 'to' => 'G2'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'G2', 'to' => 'G3'],
                    ],

                    'merge_cells' => ['C3:D3'],

                    'copy_width' => [
                        ['from' => 'E', 'to' => 'F'],
                        ['from' => 'E', 'to' => 'G'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'E1', 'to' => 'F1'],
                        ['from' => 'E1', 'to' => 'G1'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'F2'],
                        ['from' => 'E2', 'to' => 'G2'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'G2', 'to' => 'G3'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_combination2_limit()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
                        'E' => '[columns.1.one.0]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.1.b]',
                        'C' => '[rows.1.c]',
                        'E' => '[columns.1.two.0]',
                    ],
                    3 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'rows' => [
                        ['b' => 'b1', 'c' => 'c1'],
                        ['b' => 'b2', 'c' => 'c2'],
                    ],
                    'columns' => [
                        [
                            'one' => ['01', '02', '03'],
                            'two' => [15000, 20000, 30000],
                        ],
                    ],
                    'foo' => 'test1',
                    'bar' => 'test2',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'test1',
                            'E' => '01',
                            'F' => '02',
                            'G' => '03',
                        ],
                        2 => [
                            'B' => 'b2',
                            'C' => 'c2',
                            'E' => 15000,
                            'F' => 20000,
                            'G' => 30000,
                        ],
                        3 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [
                        ['from' => 'E1', 'to' => 'F1'],
                        ['from' => 'E1', 'to' => 'G1'],
                        ['from' => 'E2', 'to' => 'F2'],
                        ['from' => 'E2', 'to' => 'G2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'E', 'to' => 'F'],
                        ['from' => 'E', 'to' => 'G'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'E1', 'to' => 'F1'],
                        ['from' => 'E1', 'to' => 'G1'],
                        ['from' => 'E2', 'to' => 'F2'],
                        ['from' => 'E2', 'to' => 'G2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_combination3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'B' => '[title] [=column.month]',
                        'C' => '=D6 [oops]',
                        'D' => 'D6',
                    ],
                    3 => [
                        'G' => '[column.month] [= column.month]',
                    ],
                    4 => [
                        'A' => '[list.name]',
                        'B' => '[list.count]',
                        'C' => 'kg',
                        'D' => '[list.price]',
                        'E' => '[comment]',
                        'G' => '[column.amount]',
                        'Q' => '=A1+B3+C4+D6',
                    ],
                    6 => [
                        'B' => '[total.count]',
                    ],
                    7 => [
                        'C' => '=D6 [oops]',
                        'D' => 'D6',
                    ],
                ],

                'data' => [
                    'title' => 'foo',
                    'total' => ['count' => 3],
                    'comment' => 'bar',
                    'list' => [
                        ['name' => 'Product 1', 'count' => 2, 'price' => 753.14],
                        ['name' => 'Product 2', 'count' => 1, 'price' => 123],
                    ],
                    'column' => [
                        [
                            'month' => ['01', '02', '03'],
                            'amount' => [15000, 20000, 30000],
                        ],
                    ],
                ],

                'merge_cells' => ['G4:H4'],
            ],

            [
                'values' => [
                    1 => [
                        'B' => '[title] [=column.*.month]',
                        'C' => '=D6',
                        'D' => 'D6',
                    ],
                    3 => [
                        'G' => '[column.*.month] [= column.*.month.*]',
                    ],
                    4 => [
                        'A' => '[list.*.name]',
                        'B' => '[list.*.count]',
                        'C' => 'kg',
                        'D' => '[list.price]',
                        'E' => '[comment]',
                        'G' => '[column.*.amount.*]',
                        'Q' => '=A1+B3+C4+D6',
                    ],
                    6 => [
                        'B' => '[total.count]',
                    ],
                    7 => [
                        'C' => '=D6 [oops]',
                        'D' => 'D6',
                    ],
                ],

                'data' => [
                    'title' => 'foo',
                    'total' => ['count' => 3],
                    'comment' => 'bar',
                    'list' => [
                        ['name' => 'Product 1', 'count' => 2, 'price' => 753.14],
                        ['name' => 'Product 2', 'count' => 1, 'price' => 123],
                    ],
                    'column' => [
                        [
                            'month' => ['01', '02', '03'],
                            'amount' => [15000, 20000, 30000],
                        ],
                    ],
                ],

                'merge_cells' => ['G4:H4'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'foo',
                            'C' => '=D7',
                        ],
                        3 => [
                            'G' => '01',
                            'H' => '02',
                            'I' => '03',
                        ],
                        4 => [
                            'A' => 'Product 1',
                            'B' => 2,
                            'C' => 'kg',
                            'D' => 753.14,
                            'E' => 'bar',
                            'G' => 15000,
                            'I' => 20000,
                            'K' => 30000,
                            'Q' => '=A1+B3+C4+D7',
                        ],
                        5 => [
                            'A' => 'Product 2',
                            'B' => 1,
                            'C' => 'kg',
                            'D' => 123,
                            'E' => 'bar',
                            'G' => null,
                            'I' => null,
                            'K' => null,
                            'Q' => '=A1+B3+C5+D7',
                        ],
                        7 => [
                            'B' => 3,
                        ],
                        8 => [
                            'C' => '=D7',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 5],
                    ],

                    'copy_style' => [
                        ['from' => 'A4', 'to' => 'A5'],
                        ['from' => 'B4', 'to' => 'B5'],
                        ['from' => 'C4', 'to' => 'C5'],
                        ['from' => 'D4', 'to' => 'D5'],
                        ['from' => 'E4', 'to' => 'E5'],
                        ['from' => 'G3', 'to' => 'H3'],
                        ['from' => 'G3', 'to' => 'I3'],
                        ['from' => 'G4', 'to' => 'G5'],
                        ['from' => 'G4', 'to' => 'I4'],
                        ['from' => 'G4', 'to' => 'K4'],
                        ['from' => 'H4', 'to' => 'H5'],
                        ['from' => 'H4', 'to' => 'J4'],
                        ['from' => 'H4', 'to' => 'L4'],
                        ['from' => 'I4', 'to' => 'I5'],
                        ['from' => 'J4', 'to' => 'J5'],
                        ['from' => 'K4', 'to' => 'K5'],
                        ['from' => 'L4', 'to' => 'L5'],
                        ['from' => 'Q4', 'to' => 'Q5'],
                    ],

                    'merge_cells' => [
                        'I4:J4', 'K4:L4',
                        'G5:H5', 'I5:J5', 'K5:L5',
                    ],

                    'copy_width' => [
                        ['from' => 'G', 'to' => 'H'],
                        ['from' => 'G', 'to' => 'I'],
                        ['from' => 'G', 'to' => 'K'],
                        ['from' => 'H', 'to' => 'J'],
                        ['from' => 'H', 'to' => 'L'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A4', 'to' => 'A5'],
                        ['from' => 'B4', 'to' => 'B5'],
                        ['from' => 'C4', 'to' => 'C5'],
                        ['from' => 'D4', 'to' => 'D5'],
                        ['from' => 'E4', 'to' => 'E5'],
                        ['from' => 'G3', 'to' => 'H3'],
                        ['from' => 'G3', 'to' => 'I3'],
                        ['from' => 'G4', 'to' => 'G5'],
                        ['from' => 'G4', 'to' => 'I4'],
                        ['from' => 'G4', 'to' => 'K4'],
                        ['from' => 'I4', 'to' => 'I5'],
                        ['from' => 'K4', 'to' => 'K5'],
                        ['from' => 'Q4', 'to' => 'Q5'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_combination3_limit()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'B' => '[title]',
                    ],
                    3 => [
                        'G' => '[column.0.month.0] [= column.0.month]',
                    ],
                    4 => [
                        'A' => '[list.0.name] [=column.0.month.0]',
                        'B' => '[list.0.count]',
                        'C' => 'kg',
                        'D' => '[list.0.price]',
                        'E' => '[comment]',
                        'G' => '[column.0.amount.0]',
                    ],
                    6 => [
                        'B' => '[total.count]',
                    ],
                ],

                'data' => [
                    'title' => 'foo',
                    'total' => ['count' => 3],
                    'comment' => 'bar',
                    'list' => [
                        ['name' => 'Product 1', 'count' => 2, 'price' => 753.14],
                        ['name' => 'Product 2', 'count' => 1, 'price' => 123],
                    ],
                    'column' => [
                        [
                            'month' => ['01', '02', '03'],
                            'amount' => [15000, 20000, 30000],
                        ],
                    ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'foo',
                        ],
                        3 => [
                            'G' => '01',
                            'H' => '02',
                            'I' => '03',
                        ],
                        4 => [
                            'A' => 'Product 1',
                            'B' => 2,
                            'D' => 753.14,
                            'E' => 'bar',
                            'G' => 15000,
                        ],
                        6 => [
                            'B' => 3,
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [
                        ['from' => 'G3', 'to' => 'H3'],
                        ['from' => 'G3', 'to' => 'I3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'G', 'to' => 'H'],
                        ['from' => 'G', 'to' => 'I'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'G3', 'to' => 'H3'],
                        ['from' => 'G3', 'to' => 'I3'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_table_with_formula1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=A4',
                        'B' => '=A3',
                        'C' => '=A2',
                        'D' => '=A1',
                        'E' => '=A2:A2',
                        'F' => '=B2:G2',
                    ],
                    2 => [
                        'A' => '[table.price]',
                        'B' => '[table.count]',
                        'C' => '=A2*B2',
                        'D' => '=A1+A3+A4',
                        'E' => 'A1+A3+A4',
                    ],
                    3 => [
                        'A' => '=A4',
                        'B' => '=A3',
                        'C' => '=A2',
                        'D' => '=A1',
                        'E' => '=A2:A2',
                        'F' => '=B2:G2',
                    ],
                ],

                'data' => [
                    'table' => [
                        ['price' => 11, 'count' => 1],
                        ['price' => 12, 'count' => 2],
                        ['price' => 13, 'count' => 3],
                        ['price' => 14, 'count' => 4],
                    ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=A7',
                            'B' => '=A6',
                            'C' => '=A2',
                            'D' => '=A1',
                            'E' => '=A2:A5',
                            'F' => '=B2:G5',
                        ],
                        2 => [
                            'A' => 11,
                            'B' => 1,
                            'C' => '=A2*B2',
                            'D' => '=A1+A6+A7',
                            'E' => 'A1+A3+A4',
                        ],
                        3 => [
                            'A' => 12,
                            'B' => 2,
                            'C' => '=A3*B3',
                            'D' => '=A1+A6+A7',
                            'E' => 'A1+A3+A4',
                        ],
                        4 => [
                            'A' => 13,
                            'B' => 3,
                            'C' => '=A4*B4',
                            'D' => '=A1+A6+A7',
                            'E' => 'A1+A3+A4',
                        ],
                        5 => [
                            'A' => 14,
                            'B' => 4,
                            'C' => '=A5*B5',
                            'D' => '=A1+A6+A7',
                            'E' => 'A1+A3+A4',
                        ],
                        6 => [
                            'A' => '=A7',
                            'B' => '=A6',
                            'C' => '=A2',
                            'D' => '=A1',
                            'E' => '=A2:A5',
                            'F' => '=B2:G5',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'E2', 'to' => 'E5'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'E2', 'to' => 'E5'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_table_with_formula2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=A1',
                        'B' => '=A2',
                        'C' => '=A3',
                        'D' => '=A4',
                        'E' => '=A5',
                        'F' => '=A6',
                        'G' => '=A7',
                        'H' => '=A8',
                        'J' => '=A3:A3',
                        'K' => '=A3:C3',
                    ],
                    2 => [
                        'A' => '[= hello]',
                    ],
                    3 => [
                        'A' => '[table.price]',
                        'B' => '[table.count]',
                        'C' => '=A3*B3',
                    ],
                    4 => [
                        'A' => '[=hello]',
                    ],
                    5 => [
                        'A' => '[!table]',
                    ],
                    6 => [
                        'A' => '[!table.count]',
                    ],
                    7 => [
                        'A' => '=A1',
                        'B' => '=A2',
                        'C' => '=A3',
                        'D' => '=A4',
                        'E' => '=A5',
                        'F' => '=A6',
                        'G' => '=A7',
                        'H' => '=A8',
                        'J' => '=A3:A3',
                        'K' => '=A3:C3',
                    ],
                ],

                'data' => [
                    'table' => [
                        ['price' => 11, 'count' => 1],
                        ['price' => 12, 'count' => 2],
                        ['price' => 13, 'count' => 3],
                        ['price' => 14, 'count' => 4],
                        ['price' => 15, 'count' => 5],
                    ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=A1',
                            'B' => '=A2',
                            'C' => '=A2',
                            'D' => '=A7',
                            'E' => '=A7',
                            'F' => '=A7',
                            'G' => '=A7',
                            'H' => '=A8',
                            'J' => '=A2:A6',
                            'K' => '=A2:C6',
                        ],
                        2 => [
                            'A' => 11,
                            'B' => 1,
                            'C' => '=A2*B2',
                        ],
                        3 => [
                            'A' => 12,
                            'B' => 2,
                            'C' => '=A3*B3',
                        ],
                        4 => [
                            'A' => 13,
                            'B' => 3,
                            'C' => '=A4*B4',
                        ],
                        5 => [
                            'A' => 14,
                            'B' => 4,
                            'C' => '=A5*B5',
                        ],
                        6 => [
                            'A' => 15,
                            'B' => 5,
                            'C' => '=A6*B6',
                        ],
                        7 => [
                            'A' => '=A1',
                            'B' => '=A2',
                            'C' => '=A2',
                            'D' => '=A7',
                            'E' => '=A7',
                            'F' => '=A7',
                            'G' => '=A7',
                            'H' => '=A8',
                            'J' => '=A2:A6',
                            'K' => '=A2:C6',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 2],
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'delete', 'row' => 7],
                        ['action' => 'delete', 'row' => 7],
                        ['action' => 'delete', 'row' => 7],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C6'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C6'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_table_with_formula3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=AA1',
                        'B' => '=BB2',
                        'C' => '=CC3',
                        'D' => '=DD4',
                        'E' => '=EE5',
                        'F' => '=FF6',
                        'G' => '=A3:A3',
                    ],
                    2 => [
                        'A' => '[= hello]',
                    ],
                    3 => [
                        'A' => '[table.price]',
                        'B' => '[table.count]',
                        'C' => '=A3*B3+C1+D2+E4+F5',
                    ],
                    4 => [
                        'A' => '=AA1',
                        'B' => '=BB2',
                        'C' => '=CC3',
                        'D' => '=DD4',
                        'E' => '=EE5',
                        'F' => '=FF6',
                        'G' => '=A3:A3',
                    ],
                ],

                'data' => [
                    'table' => [
                        ['price' => 11, 'count' => 1],
                        ['price' => 12, 'count' => 2],
                        ['price' => 13, 'count' => 3],
                        ['price' => 14, 'count' => 4],
                        ['price' => 15, 'count' => 5],
                    ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=AA1',
                            'B' => '=BB2',
                            'C' => '=CC2',
                            'D' => '=DD7',
                            'E' => '=EE8',
                            'F' => '=FF9',
                            'G' => '=A2:A6',
                        ],
                        2 => [
                            'A' => 11,
                            'B' => 1,
                            'C' => '=A2*B2+C1+D2+E7+F8',
                        ],
                        3 => [
                            'A' => 12,
                            'B' => 2,
                            'C' => '=A3*B3+C1+D3+E7+F8',
                        ],
                        4 => [
                            'A' => 13,
                            'B' => 3,
                            'C' => '=A4*B4+C1+D4+E7+F8',
                        ],
                        5 => [
                            'A' => 14,
                            'B' => 4,
                            'C' => '=A5*B5+C1+D5+E7+F8',
                        ],
                        6 => [
                            'A' => 15,
                            'B' => 5,
                            'C' => '=A6*B6+C1+D6+E7+F8',
                        ],
                        7 => [
                            'A' => '=AA1',
                            'B' => '=BB2',
                            'C' => '=CC2',
                            'D' => '=DD7',
                            'E' => '=EE8',
                            'F' => '=FF9',
                            'G' => '=A2:A6',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 2],
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C6'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C6'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_1x1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                    ],
                    2 => [
                        'A' => '[matrix]',
                    ],
                    3 => [
                        'A' => '[bar]',
                    ],
                ],

                'data' => [
                    'foo' => 'test1',
                    'bar' => 'test2',
                    'matrix' => [
                        [['one', 'two', 'three', 'four']],
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'test1',
                        ],
                        2 => [
                            'A' => 'one',
                            'B' => 'two',
                            'C' => 'three',
                            'D' => 'four',
                        ],
                        3 => [
                            'A' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'B2'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'D2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'A', 'to' => 'B'],
                        ['from' => 'A', 'to' => 'C'],
                        ['from' => 'A', 'to' => 'D'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'B2'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'D2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_1x2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                    ],
                    2 => [
                        'A' => '[matrix]',
                    ],
                    3 => [
                        'A' => '[bar]',
                    ],
                ],

                'data' => [
                    'foo' => 'test1',
                    'bar' => 'test2',
                    'matrix' => [
                        [['one', 'two', 'three', 'four']],
                    ],
                ],

                'merge_cells' => ['A2:B2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'test1',
                        ],
                        2 => [
                            'A' => 'one',
                            'C' => 'two',
                            'E' => 'three',
                            'G' => 'four',
                        ],
                        3 => [
                            'A' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'E2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'B2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'F2'],
                        ['from' => 'B2', 'to' => 'H2'],
                    ],

                    'merge_cells' => ['C2:D2', 'E2:F2', 'G2:H2'],

                    'copy_width' => [
                        ['from' => 'A', 'to' => 'C'],
                        ['from' => 'A', 'to' => 'E'],
                        ['from' => 'A', 'to' => 'G'],
                        ['from' => 'B', 'to' => 'D'],
                        ['from' => 'B', 'to' => 'F'],
                        ['from' => 'B', 'to' => 'H'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'E2'],
                        ['from' => 'A2', 'to' => 'G2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_1x3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                    ],
                    2 => [
                        'A' => '[matrix]',
                    ],
                    3 => [
                        'A' => '[bar]',
                    ],
                ],

                'data' => [
                    'foo' => 'test1',
                    'bar' => 'test2',
                    'matrix' => [
                        [['one', 'two', 'three', 'four']],
                    ],
                ],

                'merge_cells' => ['A2:C2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'test1',
                        ],
                        2 => [
                            'A' => 'one',
                            'D' => 'two',
                            'G' => 'three',
                            'J' => 'four',
                        ],
                        3 => [
                            'A' => 'test2',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'A2', 'to' => 'J2'],
                        ['from' => 'B2', 'to' => 'E2'],
                        ['from' => 'B2', 'to' => 'H2'],
                        ['from' => 'B2', 'to' => 'K2'],
                        ['from' => 'C2', 'to' => 'F2'],
                        ['from' => 'C2', 'to' => 'I2'],
                        ['from' => 'C2', 'to' => 'L2'],
                    ],

                    'merge_cells' => ['D2:F2', 'G2:I2', 'J2:L2'],

                    'copy_width' => [
                        ['from' => 'A', 'to' => 'D'],
                        ['from' => 'A', 'to' => 'G'],
                        ['from' => 'A', 'to' => 'J'],
                        ['from' => 'B', 'to' => 'E'],
                        ['from' => 'B', 'to' => 'H'],
                        ['from' => 'B', 'to' => 'K'],
                        ['from' => 'C', 'to' => 'F'],
                        ['from' => 'C', 'to' => 'I'],
                        ['from' => 'C', 'to' => 'L'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'A2', 'to' => 'J2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_3x1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                    ],
                    2 => [
                        'A' => '[matrix]',
                    ],
                    3 => [
                        'A' => '[bar]',
                    ],
                ],

                'data' => [
                    'foo' => 'test1',
                    'bar' => 'test2',
                    'matrix' => [
                        [['one', 'two', 'three', 'four']],
                        [['one-b', 'two-b', 'three-b', 'four-b']],
                        [['one-c', 'two-c', 'three-c', 'four-c']],
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'test1',
                        ],
                        2 => [
                            'A' => 'one',
                            'B' => 'two',
                            'C' => 'three',
                            'D' => 'four',
                        ],
                        3 => [
                            'A' => 'one-b',
                            'B' => 'two-b',
                            'C' => 'three-b',
                            'D' => 'four-b',
                        ],
                        4 => [
                            'A' => 'one-c',
                            'B' => 'two-c',
                            'C' => 'three-c',
                            'D' => 'four-c',
                        ],
                        5 => [
                            'A' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'B2'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'A', 'to' => 'B'],
                        ['from' => 'A', 'to' => 'C'],
                        ['from' => 'A', 'to' => 'D'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'B2'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_3x2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                    ],
                    2 => [
                        'A' => '[matrix]',
                    ],
                    3 => [
                        'A' => '[bar]',
                    ],
                ],

                'data' => [
                    'foo' => 'test1',
                    'bar' => 'test2',
                    'matrix' => [
                        [['one', 'two', 'three', 'four']],
                        [['one-b', 'two-b', 'three-b', 'four-b']],
                        [['one-c', 'two-c', 'three-c', 'four-c']],
                    ],
                ],

                'merge_cells' => ['A2:B2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'test1',
                        ],
                        2 => [
                            'A' => 'one',
                            'C' => 'two',
                            'E' => 'three',
                            'G' => 'four',
                        ],
                        3 => [
                            'A' => 'one-b',
                            'C' => 'two-b',
                            'E' => 'three-b',
                            'G' => 'four-b',
                        ],
                        4 => [
                            'A' => 'one-c',
                            'C' => 'two-c',
                            'E' => 'three-c',
                            'G' => 'four-c',
                        ],
                        5 => [
                            'A' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'E2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'F2'],
                        ['from' => 'B2', 'to' => 'H2'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'F2', 'to' => 'F4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'H2', 'to' => 'H4'],
                    ],

                    'merge_cells' => [
                        'C2:D2', 'E2:F2', 'G2:H2',
                        'A3:B3', 'C3:D3', 'E3:F3', 'G3:H3',
                        'A4:B4', 'C4:D4', 'E4:F4', 'G4:H4',
                    ],

                    'copy_width' => [
                        ['from' => 'A', 'to' => 'C'],
                        ['from' => 'A', 'to' => 'E'],
                        ['from' => 'A', 'to' => 'G'],
                        ['from' => 'B', 'to' => 'D'],
                        ['from' => 'B', 'to' => 'F'],
                        ['from' => 'B', 'to' => 'H'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'E2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_3x3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                    ],
                    2 => [
                        'A' => '[matrix]',
                    ],
                    3 => [
                        'A' => '[bar]',
                    ],
                ],

                'data' => [
                    'foo' => 'test1',
                    'bar' => 'test2',
                    'matrix' => [
                        [['one', 'two', 'three', 'four']],
                        [['one-b', 'two-b', 'three-b', 'four-b']],
                        [['one-c', 'two-c', 'three-c', 'four-c']],
                    ],
                ],

                'merge_cells' => ['A2:C2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'test1',
                        ],
                        2 => [
                            'A' => 'one',
                            'D' => 'two',
                            'G' => 'three',
                            'J' => 'four',
                        ],
                        3 => [
                            'A' => 'one-b',
                            'D' => 'two-b',
                            'G' => 'three-b',
                            'J' => 'four-b',
                        ],
                        4 => [
                            'A' => 'one-c',
                            'D' => 'two-c',
                            'G' => 'three-c',
                            'J' => 'four-c',
                        ],
                        5 => [
                            'A' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'A2', 'to' => 'J2'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'E2'],
                        ['from' => 'B2', 'to' => 'H2'],
                        ['from' => 'B2', 'to' => 'K2'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'F2'],
                        ['from' => 'C2', 'to' => 'I2'],
                        ['from' => 'C2', 'to' => 'L2'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'E2', 'to' => 'E4'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'F2', 'to' => 'F4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                        ['from' => 'H2', 'to' => 'H3'],
                        ['from' => 'H2', 'to' => 'H4'],
                        ['from' => 'I2', 'to' => 'I3'],
                        ['from' => 'I2', 'to' => 'I4'],
                        ['from' => 'J2', 'to' => 'J3'],
                        ['from' => 'J2', 'to' => 'J4'],
                        ['from' => 'K2', 'to' => 'K3'],
                        ['from' => 'K2', 'to' => 'K4'],
                        ['from' => 'L2', 'to' => 'L3'],
                        ['from' => 'L2', 'to' => 'L4'],
                    ],

                    'merge_cells' => [
                        'D2:F2', 'G2:I2', 'J2:L2',
                        'A3:C3', 'D3:F3', 'G3:I3', 'J3:L3',
                        'A4:C4', 'D4:F4', 'G4:I4', 'J4:L4',
                    ],

                    'copy_width' => [
                        ['from' => 'A', 'to' => 'D'],
                        ['from' => 'A', 'to' => 'G'],
                        ['from' => 'A', 'to' => 'J'],
                        ['from' => 'B', 'to' => 'E'],
                        ['from' => 'B', 'to' => 'H'],
                        ['from' => 'B', 'to' => 'K'],
                        ['from' => 'C', 'to' => 'F'],
                        ['from' => 'C', 'to' => 'I'],
                        ['from' => 'C', 'to' => 'L'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'A2', 'to' => 'J2'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'D2', 'to' => 'D4'],
                        ['from' => 'G2', 'to' => 'G3'],
                        ['from' => 'G2', 'to' => 'G4'],
                        ['from' => 'J2', 'to' => 'J3'],
                        ['from' => 'J2', 'to' => 'J4'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_multi1()
    {
        $data = [
            [
                'values' => [
                    2 => [
                        'C' => '[project.months]',
                    ],
                    3 => [
                        'A' => '[project.name]',
                        'C' => '[project.amount]',
                    ],
                ],

                'data' => [
                    'project' => [
                        'name' => ['N1', 'N2', 'N3'],
                        'months' => [['01', '02', '03']],
                        'amount' => [
                            [101, 201, 301],
                            [102, 202, 302],
                            [103, 203, 303],
                        ],
                    ],
                ],

                'merge_cells' => ['C2:D2', 'A3:B3', 'C3:D3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        2 => [
                            'C' => '01',
                            'E' => '02',
                            'G' => '03',
                        ],
                        3 => [
                            'A' => 'N1',
                            'C' => 101,
                            'E' => 201,
                            'G' => 301,
                        ],
                        4 => [
                            'A' => 'N2',
                            'C' => 102,
                            'E' => 202,
                            'G' => 302,
                        ],
                        5 => [
                            'A' => 'N3',
                            'C' => 103,
                            'E' => 203,
                            'G' => 303,
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                    ],

                    'copy_style' => [
                        ['from' => 'A3', 'to' => 'A4'],
                        ['from' => 'A3', 'to' => 'A5'],
                        ['from' => 'B3', 'to' => 'B4'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'C2', 'to' => 'G2'],
                        ['from' => 'C3', 'to' => 'C4'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'E3'],
                        ['from' => 'C3', 'to' => 'G3'],
                        ['from' => 'D2', 'to' => 'F2'],
                        ['from' => 'D2', 'to' => 'H2'],
                        ['from' => 'D3', 'to' => 'D4'],
                        ['from' => 'D3', 'to' => 'D5'],
                        ['from' => 'D3', 'to' => 'F3'],
                        ['from' => 'D3', 'to' => 'H3'],
                        ['from' => 'E3', 'to' => 'E4'],
                        ['from' => 'E3', 'to' => 'E5'],
                        ['from' => 'F3', 'to' => 'F4'],
                        ['from' => 'F3', 'to' => 'F5'],
                        ['from' => 'G3', 'to' => 'G4'],
                        ['from' => 'G3', 'to' => 'G5'],
                        ['from' => 'H3', 'to' => 'H4'],
                        ['from' => 'H3', 'to' => 'H5'],
                    ],

                    'merge_cells' => [
                        'E2:F2', 'G2:H2',
                        'E3:F3', 'G3:H3',
                        'A4:B4', 'C4:D4', 'E4:F4', 'G4:H4',
                        'A5:B5', 'C5:D5', 'E5:F5', 'G5:H5',
                    ],

                    'copy_width' => [
                        ['from' => 'C', 'to' => 'E'],
                        ['from' => 'C', 'to' => 'G'],
                        ['from' => 'D', 'to' => 'F'],
                        ['from' => 'D', 'to' => 'H'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A3', 'to' => 'A4'],
                        ['from' => 'A3', 'to' => 'A5'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'C2', 'to' => 'G2'],
                        ['from' => 'C3', 'to' => 'C4'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'E3'],
                        ['from' => 'C3', 'to' => 'G3'],
                        ['from' => 'E3', 'to' => 'E4'],
                        ['from' => 'E3', 'to' => 'E5'],
                        ['from' => 'G3', 'to' => 'G4'],
                        ['from' => 'G3', 'to' => 'G5'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_matrix_merge_multi2()
    {
        $data = [
            [
                'values' => [
                    2 => [
                        'D' => '[project.months]',
                    ],
                    3 => [
                        'A' => '[project.name]',
                        'D' => '[project.amount]',
                    ],
                ],

                'data' => [
                    'project' => [
                        'name' => ['N1', 'N2', 'N3'],
                        'months' => [['01', '02', '03']],
                        'amount' => [
                            [101, 201, 301],
                            [102, 202, 302],
                            [103, 203, 303],
                        ],
                    ],
                ],

                'merge_cells' => ['D2:F2', 'A3:C3', 'D3:F3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        2 => [
                            'D' => '01',
                            'G' => '02',
                            'J' => '03',
                        ],
                        3 => [
                            'A' => 'N1',
                            'D' => 101,
                            'G' => 201,
                            'J' => 301,
                        ],
                        4 => [
                            'A' => 'N2',
                            'D' => 102,
                            'G' => 202,
                            'J' => 302,
                        ],
                        5 => [
                            'A' => 'N3',
                            'D' => 103,
                            'G' => 203,
                            'J' => 303,
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                    ],

                    'copy_style' => [
                        ['from' => 'A3', 'to' => 'A4'],
                        ['from' => 'A3', 'to' => 'A5'],
                        ['from' => 'B3', 'to' => 'B4'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'C3', 'to' => 'C4'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'D2', 'to' => 'G2'],
                        ['from' => 'D2', 'to' => 'J2'],
                        ['from' => 'D3', 'to' => 'D4'],
                        ['from' => 'D3', 'to' => 'D5'],
                        ['from' => 'D3', 'to' => 'G3'],
                        ['from' => 'D3', 'to' => 'J3'],
                        ['from' => 'E2', 'to' => 'H2'],
                        ['from' => 'E2', 'to' => 'K2'],
                        ['from' => 'E3', 'to' => 'E4'],
                        ['from' => 'E3', 'to' => 'E5'],
                        ['from' => 'E3', 'to' => 'H3'],
                        ['from' => 'E3', 'to' => 'K3'],
                        ['from' => 'F2', 'to' => 'I2'],
                        ['from' => 'F2', 'to' => 'L2'],
                        ['from' => 'F3', 'to' => 'F4'],
                        ['from' => 'F3', 'to' => 'F5'],
                        ['from' => 'F3', 'to' => 'I3'],
                        ['from' => 'F3', 'to' => 'L3'],
                        ['from' => 'G3', 'to' => 'G4'],
                        ['from' => 'G3', 'to' => 'G5'],
                        ['from' => 'H3', 'to' => 'H4'],
                        ['from' => 'H3', 'to' => 'H5'],
                        ['from' => 'I3', 'to' => 'I4'],
                        ['from' => 'I3', 'to' => 'I5'],
                        ['from' => 'J3', 'to' => 'J4'],
                        ['from' => 'J3', 'to' => 'J5'],
                        ['from' => 'K3', 'to' => 'K4'],
                        ['from' => 'K3', 'to' => 'K5'],
                        ['from' => 'L3', 'to' => 'L4'],
                        ['from' => 'L3', 'to' => 'L5'],
                    ],

                    'merge_cells' => [
                        'G2:I2', 'J2:L2',
                        'G3:I3', 'J3:L3',
                        'A4:C4', 'D4:F4', 'G4:I4', 'J4:L4',
                        'A5:C5', 'D5:F5', 'G5:I5', 'J5:L5',
                    ],

                    'copy_width' => [
                        ['from' => 'D', 'to' => 'G'],
                        ['from' => 'D', 'to' => 'J'],
                        ['from' => 'E', 'to' => 'H'],
                        ['from' => 'E', 'to' => 'K'],
                        ['from' => 'F', 'to' => 'I'],
                        ['from' => 'F', 'to' => 'L'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A3', 'to' => 'A4'],
                        ['from' => 'A3', 'to' => 'A5'],
                        ['from' => 'D2', 'to' => 'G2'],
                        ['from' => 'D2', 'to' => 'J2'],
                        ['from' => 'D3', 'to' => 'D4'],
                        ['from' => 'D3', 'to' => 'D5'],
                        ['from' => 'D3', 'to' => 'G3'],
                        ['from' => 'D3', 'to' => 'J3'],
                        ['from' => 'G3', 'to' => 'G4'],
                        ['from' => 'G3', 'to' => 'G5'],
                        ['from' => 'J3', 'to' => 'J4'],
                        ['from' => 'J3', 'to' => 'J5'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_0x2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                        ],
                        2 => [
                            'A' => null,
                            'B' => null,
                        ],
                        3 => [
                            'B' => null,
                            'C' => null,
                        ],
                        4 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_1x2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['id' => 1, 'name' => 'project 1', 'amount_1' => 101, 'amount_2' => 102],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                        ],
                        2 => [
                            'A' => 1,
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 101,
                            'C' => 102,
                        ],
                        4 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_2x2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['id' => 1, 'name' => 'project 1', 'amount_1' => 101, 'amount_2' => 102],
                        ['id' => 2, 'name' => 'project 2', 'amount_1' => 201, 'amount_2' => 202],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                        ],
                        2 => [
                            'A' => 1,
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 101,
                            'C' => 102,
                        ],
                        4 => [
                            'A' => 2,
                            'B' => 'project 2',
                        ],
                        5 => [
                            'B' => 201,
                            'C' => 202,
                        ],
                        6 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'C3', 'to' => 'C5'],
                    ],

                    'merge_cells' => [
                        'A4:A5',
                    ],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'C3', 'to' => 'C5'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_2x2_skip()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B2:B2)',
                    ],
                    2 => [
                        'A' => 'Hello',
                        'B' => '[project.name]',
                    ],
                    4 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['name' => 'project 1'],
                        ['name' => 'project 2'],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A3)',
                            'B' => '=SUM(B2:B3)',
                        ],
                        2 => [
                            'A' => 'Hello',
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 'project 2',
                        ],
                        5 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                    ],

                    'copy_style' => [
                        ['from' => 'B2', 'to' => 'B3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'B2', 'to' => 'B3'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_3x2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['id' => 1, 'name' => 'project 1', 'amount_1' => 101, 'amount_2' => 102],
                        ['id' => 2, 'name' => 'project 2', 'amount_1' => 201, 'amount_2' => 202],
                        ['id' => 3, 'name' => 'project 3', 'amount_1' => 301, 'amount_2' => 302],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                        ],
                        2 => [
                            'A' => 1,
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 101,
                            'C' => 102,
                        ],
                        4 => [
                            'A' => 2,
                            'B' => 'project 2',
                        ],
                        5 => [
                            'B' => 201,
                            'C' => 202,
                        ],
                        6 => [
                            'A' => 3,
                            'B' => 'project 3',
                        ],
                        7 => [
                            'B' => 301,
                            'C' => 302,
                        ],
                        8 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'add', 'row' => 7],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'B3', 'to' => 'B7'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'C7'],
                    ],

                    'merge_cells' => [
                        'A4:A5', 'A6:A7',
                    ],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'B3', 'to' => 'B7'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'C7'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_4x2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['id' => 1, 'name' => 'project 1', 'amount_1' => 101, 'amount_2' => 102],
                        ['id' => 2, 'name' => 'project 2', 'amount_1' => 201, 'amount_2' => 202],
                        ['id' => 3, 'name' => 'project 3', 'amount_1' => 301, 'amount_2' => 302],
                        ['id' => 4, 'name' => 'project 4', 'amount_1' => 401, 'amount_2' => 402],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                        ],
                        2 => [
                            'A' => 1,
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 101,
                            'C' => 102,
                        ],
                        4 => [
                            'A' => 2,
                            'B' => 'project 2',
                        ],
                        5 => [
                            'B' => 201,
                            'C' => 202,
                        ],
                        6 => [
                            'A' => 3,
                            'B' => 'project 3',
                        ],
                        7 => [
                            'B' => 301,
                            'C' => 302,
                        ],
                        8 => [
                            'A' => 4,
                            'B' => 'project 4',
                        ],
                        9 => [
                            'B' => 401,
                            'C' => 402,
                        ],
                        10 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'add', 'row' => 7],
                        ['action' => 'add', 'row' => 8],
                        ['action' => 'add', 'row' => 9],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'B3', 'to' => 'B7'],
                        ['from' => 'B3', 'to' => 'B9'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'C7'],
                        ['from' => 'C3', 'to' => 'C9'],
                    ],

                    'merge_cells' => [
                        'A4:A5', 'A6:A7', 'A8:A9',
                    ],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'B3', 'to' => 'B7'],
                        ['from' => 'B3', 'to' => 'B9'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'C7'],
                        ['from' => 'C3', 'to' => 'C9'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_4x3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                        'C' => '=SUM(C4:C4)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                    ],
                    4 => [
                        'C' => '[project.amount_2]',
                    ],
                    6 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['id' => 1, 'name' => 'project 1', 'amount_1' => 101, 'amount_2' => 102],
                        ['id' => 2, 'name' => 'project 2', 'amount_1' => 201, 'amount_2' => 202],
                        ['id' => 3, 'name' => 'project 3', 'amount_1' => 301, 'amount_2' => 302],
                        ['id' => 4, 'name' => 'project 4', 'amount_1' => 401, 'amount_2' => 402],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A4'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                            'C' => '=SUM(C4:C4)',
                        ],
                        2 => [
                            'A' => 1,
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 101,
                        ],
                        4 => [
                            'C' => 102,
                        ],
                        5 => [
                            'A' => 2,
                            'B' => 'project 2',
                        ],
                        6 => [
                            'B' => 201,
                        ],
                        7 => [
                            'C' => 202,
                        ],
                        8 => [
                            'A' => 3,
                            'B' => 'project 3',
                        ],
                        9 => [
                            'B' => 301,
                        ],
                        10 => [
                            'C' => 302,
                        ],
                        11 => [
                            'A' => 4,
                            'B' => 'project 4',
                        ],
                        12 => [
                            'B' => 401,
                        ],
                        13 => [
                            'C' => 402,
                        ],
                        15 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'add', 'row' => 7],
                        ['action' => 'add', 'row' => 8],
                        ['action' => 'add', 'row' => 9],
                        ['action' => 'add', 'row' => 10],
                        ['action' => 'add', 'row' => 11],
                        ['action' => 'add', 'row' => 12],
                        ['action' => 'add', 'row' => 13],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'A2', 'to' => 'A11'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B2', 'to' => 'B11'],
                        ['from' => 'B3', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B9'],
                        ['from' => 'B3', 'to' => 'B12'],
                        ['from' => 'C4', 'to' => 'C7'],
                        ['from' => 'C4', 'to' => 'C10'],
                        ['from' => 'C4', 'to' => 'C13'],
                    ],

                    'merge_cells' => [
                        'A5:A7', 'A8:A10', 'A11:A13',
                    ],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'A2', 'to' => 'A11'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B2', 'to' => 'B11'],
                        ['from' => 'B3', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B9'],
                        ['from' => 'B3', 'to' => 'B12'],
                        ['from' => 'C4', 'to' => 'C7'],
                        ['from' => 'C4', 'to' => 'C10'],
                        ['from' => 'C4', 'to' => 'C13'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_4x4()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                        'C' => '=SUM(C4:C4)',
                        'D' => '=SUM(D5:D5)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                    ],
                    4 => [
                        'C' => '[project.amount_2]',
                    ],
                    6 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['id' => 1, 'name' => 'project 1', 'amount_1' => 101, 'amount_2' => 102],
                        ['id' => 2, 'name' => 'project 2', 'amount_1' => 201, 'amount_2' => 202],
                        ['id' => 3, 'name' => 'project 3', 'amount_1' => 301, 'amount_2' => 302],
                        ['id' => 4, 'name' => 'project 4', 'amount_1' => 401, 'amount_2' => 402],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A5'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                            'C' => '=SUM(C4:C4)',
                            'D' => '=SUM(D5:D5)',
                        ],
                        2 => [
                            'A' => 1,
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 101,
                        ],
                        4 => [
                            'C' => 102,
                        ],
                        6 => [
                            'A' => 2,
                            'B' => 'project 2',
                        ],
                        7 => [
                            'B' => 201,
                        ],
                        8 => [
                            'C' => 202,
                        ],
                        10 => [
                            'A' => 3,
                            'B' => 'project 3',
                        ],
                        11 => [
                            'B' => 301,
                        ],
                        12 => [
                            'C' => 302,
                        ],
                        14 => [
                            'A' => 4,
                            'B' => 'project 4',
                        ],
                        15 => [
                            'B' => 401,
                        ],
                        16 => [
                            'C' => 402,
                        ],
                        18 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'add', 'row' => 7],
                        ['action' => 'add', 'row' => 8],
                        ['action' => 'add', 'row' => 9],
                        ['action' => 'add', 'row' => 10],
                        ['action' => 'add', 'row' => 11],
                        ['action' => 'add', 'row' => 12],
                        ['action' => 'add', 'row' => 13],
                        ['action' => 'add', 'row' => 14],
                        ['action' => 'add', 'row' => 15],
                        ['action' => 'add', 'row' => 16],
                        ['action' => 'add', 'row' => 17],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'A2', 'to' => 'A10'],
                        ['from' => 'A2', 'to' => 'A14'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'B2', 'to' => 'B10'],
                        ['from' => 'B2', 'to' => 'B14'],
                        ['from' => 'B3', 'to' => 'B7'],
                        ['from' => 'B3', 'to' => 'B11'],
                        ['from' => 'B3', 'to' => 'B15'],
                        ['from' => 'C4', 'to' => 'C8'],
                        ['from' => 'C4', 'to' => 'C12'],
                        ['from' => 'C4', 'to' => 'C16'],
                    ],

                    'merge_cells' => [
                        'A6:A9', 'A10:A13', 'A14:A17',
                    ],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'A2', 'to' => 'A10'],
                        ['from' => 'A2', 'to' => 'A14'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'B2', 'to' => 'B10'],
                        ['from' => 'B2', 'to' => 'B14'],
                        ['from' => 'B3', 'to' => 'B7'],
                        ['from' => 'B3', 'to' => 'B11'],
                        ['from' => 'B3', 'to' => 'B15'],
                        ['from' => 'C4', 'to' => 'C8'],
                        ['from' => 'C4', 'to' => 'C12'],
                        ['from' => 'C4', 'to' => 'C16'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_4x5()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(B3:B3)',
                        'C' => '=SUM(C4:C4)',
                        'D' => '=SUM(D5:D5)',
                        'E' => '=SUM(E6:E6)',
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                    ],
                    5 => [
                        'C' => '[project.amount_2]',
                    ],
                    7 => [
                        'A' => '[foo]',
                    ],
                ],

                'data' => [
                    'project' => [
                        ['id' => 1, 'name' => 'project 1', 'amount_1' => 101, 'amount_2' => 102],
                        ['id' => 2, 'name' => 'project 2', 'amount_1' => 201, 'amount_2' => 202],
                        ['id' => 3, 'name' => 'project 3', 'amount_1' => 301, 'amount_2' => 302],
                        ['id' => 4, 'name' => 'project 4', 'amount_1' => 401, 'amount_2' => 402],
                    ],
                    'foo' => 'test1',
                ],

                'merge_cells' => ['A2:A6'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A2)',
                            'B' => '=SUM(B3:B3)',
                            'C' => '=SUM(C4:C4)',
                            'D' => '=SUM(D5:D5)',
                            'E' => '=SUM(E6:E6)',
                        ],
                        2 => [
                            'A' => 1,
                            'B' => 'project 1',
                        ],
                        3 => [
                            'B' => 101,
                        ],
                        5 => [
                            'C' => 102,
                        ],
                        7 => [
                            'A' => 2,
                            'B' => 'project 2',
                        ],
                        8 => [
                            'B' => 201,
                        ],
                        10 => [
                            'C' => 202,
                        ],
                        12 => [
                            'A' => 3,
                            'B' => 'project 3',
                        ],
                        13 => [
                            'B' => 301,
                        ],
                        15 => [
                            'C' => 302,
                        ],
                        17 => [
                            'A' => 4,
                            'B' => 'project 4',
                        ],
                        18 => [
                            'B' => 401,
                        ],
                        20 => [
                            'C' => 402,
                        ],
                        22 => [
                            'A' => 'test1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 7],
                        ['action' => 'add', 'row' => 8],
                        ['action' => 'add', 'row' => 9],
                        ['action' => 'add', 'row' => 10],
                        ['action' => 'add', 'row' => 11],
                        ['action' => 'add', 'row' => 12],
                        ['action' => 'add', 'row' => 13],
                        ['action' => 'add', 'row' => 14],
                        ['action' => 'add', 'row' => 15],
                        ['action' => 'add', 'row' => 16],
                        ['action' => 'add', 'row' => 17],
                        ['action' => 'add', 'row' => 18],
                        ['action' => 'add', 'row' => 19],
                        ['action' => 'add', 'row' => 20],
                        ['action' => 'add', 'row' => 21],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A7'],
                        ['from' => 'A2', 'to' => 'A12'],
                        ['from' => 'A2', 'to' => 'A17'],
                        ['from' => 'B2', 'to' => 'B7'],
                        ['from' => 'B2', 'to' => 'B12'],
                        ['from' => 'B2', 'to' => 'B17'],
                        ['from' => 'B3', 'to' => 'B8'],
                        ['from' => 'B3', 'to' => 'B13'],
                        ['from' => 'B3', 'to' => 'B18'],
                        ['from' => 'C5', 'to' => 'C10'],
                        ['from' => 'C5', 'to' => 'C15'],
                        ['from' => 'C5', 'to' => 'C20'],
                    ],

                    'merge_cells' => [
                        'A7:A11', 'A12:A16', 'A17:A21',
                    ],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A7'],
                        ['from' => 'A2', 'to' => 'A12'],
                        ['from' => 'A2', 'to' => 'A17'],
                        ['from' => 'B2', 'to' => 'B7'],
                        ['from' => 'B2', 'to' => 'B12'],
                        ['from' => 'B2', 'to' => 'B17'],
                        ['from' => 'B3', 'to' => 'B8'],
                        ['from' => 'B3', 'to' => 'B13'],
                        ['from' => 'B3', 'to' => 'B18'],
                        ['from' => 'C5', 'to' => 'C10'],
                        ['from' => 'C5', 'to' => 'C15'],
                        ['from' => 'C5', 'to' => 'C20'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_list_merge_multi()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'B' => '[names]',
                    ],
                    2 => [
                        'A' => '[months]',
                        'B' => '[amounts.p1]',
                    ],
                    4 => [
                        'B' => '[amounts.p2]',
                    ],
                    '5' => [
                        'A' => '=A1',
                    ],
                ],

                'data' => [
                    'names' => [['One', 'Two', 'Three']],
                    'months' => ['01', '02', '03'],
                    'amounts' => [
                        'p1' => [
                            ['One R1 P1', 'Two R1 P1', 'Three R1 P1'],
                            ['One R2 P1', 'Two R2 P1', 'Three R2 P1'],
                            ['One R3 P1', 'Two R3 P1', 'Three R3 P1'],
                        ],
                        'p2' => [
                            ['One R1 P2', 'Two R1 P2', 'Three R1 P2'],
                            ['One R2 P2', 'Two R2 P2', 'Three R2 P2'],
                            ['One R3 P2', 'Two R3 P2', 'Three R3 P2'],
                        ],
                    ],
                ],

                'merge_cells' => ['A2:A4', 'B1:C1', 'B2:C2', 'B4:C4'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'One',
                            'D' => 'Two',
                            'F' => 'Three',
                        ],
                        2 => [
                            'A' => '01',
                            'B' => 'One R1 P1',
                            'D' => 'Two R1 P1',
                            'F' => 'Three R1 P1',
                        ],
                        4 => [
                            'B' => 'One R1 P2',
                            'D' => 'Two R1 P2',
                            'F' => 'Three R1 P2',
                        ],
                        5 => [
                            'A' => '02',
                            'B' => 'One R2 P1',
                            'D' => 'Two R2 P1',
                            'F' => 'Three R2 P1',
                        ],
                        7 => [
                            'B' => 'One R2 P2',
                            'D' => 'Two R2 P2',
                            'F' => 'Three R2 P2',
                        ],
                        8 => [
                            'A' => '03',
                            'B' => 'One R3 P1',
                            'D' => 'Two R3 P1',
                            'F' => 'Three R3 P1',
                        ],
                        10 => [
                            'B' => 'One R3 P2',
                            'D' => 'Two R3 P2',
                            'F' => 'Three R3 P2',
                        ],
                        11 => [
                            'A' => '=A1',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'add', 'row' => 7],
                        ['action' => 'add', 'row' => 8],
                        ['action' => 'add', 'row' => 9],
                        ['action' => 'add', 'row' => 10],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'B1', 'to' => 'D1'],
                        ['from' => 'B1', 'to' => 'F1'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'F2'],
                        ['from' => 'B4', 'to' => 'B7'],
                        ['from' => 'B4', 'to' => 'B10'],
                        ['from' => 'B4', 'to' => 'D4'],
                        ['from' => 'B4', 'to' => 'F4'],
                        ['from' => 'C1', 'to' => 'E1'],
                        ['from' => 'C1', 'to' => 'G1'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C8'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'C2', 'to' => 'G2'],
                        ['from' => 'C4', 'to' => 'C7'],
                        ['from' => 'C4', 'to' => 'C10'],
                        ['from' => 'C4', 'to' => 'E4'],
                        ['from' => 'C4', 'to' => 'G4'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'D2', 'to' => 'D8'],
                        ['from' => 'D4', 'to' => 'D7'],
                        ['from' => 'D4', 'to' => 'D10'],
                        ['from' => 'E2', 'to' => 'E5'],
                        ['from' => 'E2', 'to' => 'E8'],
                        ['from' => 'E4', 'to' => 'E7'],
                        ['from' => 'E4', 'to' => 'E10'],
                        ['from' => 'F2', 'to' => 'F5'],
                        ['from' => 'F2', 'to' => 'F8'],
                        ['from' => 'F4', 'to' => 'F7'],
                        ['from' => 'F4', 'to' => 'F10'],
                        ['from' => 'G2', 'to' => 'G5'],
                        ['from' => 'G2', 'to' => 'G8'],
                        ['from' => 'G4', 'to' => 'G7'],
                        ['from' => 'G4', 'to' => 'G10'],
                    ],

                    'merge_cells' => [
                        'D1:E1', 'F1:G1',
                        'D2:E2', 'F2:G2',
                        'A5:A7', 'B5:C5', 'D5:E5', 'F5:G5',
                        'A8:A10', 'B8:C8', 'D8:E8', 'F8:G8',
                        'D4:E4', 'F4:G4',
                        'B7:C7', 'D7:E7', 'F7:G7',
                        'B10:C10', 'D10:E10', 'F10:G10',
                    ],

                    'copy_width' => [
                        ['from' => 'B', 'to' => 'D'],
                        ['from' => 'B', 'to' => 'F'],
                        ['from' => 'C', 'to' => 'E'],
                        ['from' => 'C', 'to' => 'G'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'B1', 'to' => 'D1'],
                        ['from' => 'B1', 'to' => 'F1'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'F2'],
                        ['from' => 'B4', 'to' => 'B7'],
                        ['from' => 'B4', 'to' => 'B10'],
                        ['from' => 'B4', 'to' => 'D4'],
                        ['from' => 'B4', 'to' => 'F4'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'D2', 'to' => 'D8'],
                        ['from' => 'D4', 'to' => 'D7'],
                        ['from' => 'D4', 'to' => 'D10'],
                        ['from' => 'F2', 'to' => 'F5'],
                        ['from' => 'F2', 'to' => 'F8'],
                        ['from' => 'F4', 'to' => 'F7'],
                        ['from' => 'F4', 'to' => 'F10'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_shift1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[list.qty]',
                    ],
                    2 => [
                        'A' => '=SUM(A1:A1)',
                        'B' => '=SUM(A1:A10)',
                    ],
                    3 => [
                        'A' => '-',
                        'b' => '[! list]',
                    ],
                ],

                'data' => [
                    'list' => [ ['qty' => 1], ['qty' => 2] ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 1,
                        ],
                        2 => [
                            'A' => 2,
                        ],
                        3 => [
                            'A' => '=SUM(A1:A2)',
                            'B' => '=SUM(A1:A10)',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2],
                        ['action' => 'delete', 'row' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A1', 'to' => 'A2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_shift2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(A2:A20)',
                    ],
                    2 => [
                        'A' => '[list.qty]',
                    ],
                    3 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(A2:A20)',
                    ],
                    4 => [
                        'A' => '-',
                        'b' => '[! list]',
                    ],
                ],

                'data' => [
                    'list' => [ ['qty' => 1], ['qty' => 2] ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=SUM(A2:A3)',
                            'B' => '=SUM(A2:A20)',
                        ],
                        2 => [
                            'A' => 1,
                        ],
                        3 => [
                            'A' => 2,
                        ],
                        4 => [
                            'A' => '=SUM(A2:A3)',
                            'B' => '=SUM(A2:A20)',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'delete', 'row' => 5],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_shift3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[list.qty]',
                    ],
                    2 => [
                        'A' => '=SUM(A1:A1)',
                    ],
                ],

                'data' => [
                    'list' => [ ['qty' => 1], ['qty' => 2] ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 1,
                        ],
                        2 => [
                            'A' => 2,
                        ],
                        3 => [
                            'A' => '=SUM(A1:A2)',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A1', 'to' => 'A2'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_shift4()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=A2:A2',
                    ],
                    2 => [
                        'A' => '[list.qty]',
                    ],
                ],

                'data' => [
                    'list' => [ ['qty' => 1], ['qty' => 2] ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '=A2:A3',
                        ],
                        2 => [
                            'A' => 1,
                        ],
                        3 => [
                            'A' => 2,
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A3'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_merge1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[list.a]',
                        'B' => '[list.b]',
                        'C' => '[list.c]',
                    ],
                    2 => [
                        'C' => 'Foo',
                    ],
                ],

                'data' => [
                    'list' => [
                        ['a' => 'a1', 'b' => 'b1', 'c' => 'c1'],
                        ['a' => 'a2', 'b' => 'b2', 'c' => 'c2'],
                        ['a' => 'a3', 'b' => 'b3', 'c' => 'c3'],
                        ['a' => 'a4', 'b' => 'b4', 'c' => 'c4'],
                    ],
                ],

                'merge_cells' => ['A1:A3', 'B1:B2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'a1',
                            'B' => 'b1',
                            'C' => 'c1',
                        ],
                        2 => [
                            'C' => 'Foo',
                        ],
                        4 => [
                            'A' => 'a2',
                            'B' => 'b2',
                            'C' => 'c2',
                        ],
                        5 => [
                            'C' => 'Foo',
                        ],
                        7 => [
                            'A' => 'a3',
                            'B' => 'b3',
                            'C' => 'c3',
                        ],
                        8 => [
                            'C' => 'Foo',
                        ],
                        10 => [
                            'A' => 'a4',
                            'B' => 'b4',
                            'C' => 'c4',
                        ],
                        11 => [
                            'C' => 'Foo',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 4],
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'add', 'row' => 7],
                        ['action' => 'add', 'row' => 8],
                        ['action' => 'add', 'row' => 9],
                        ['action' => 'add', 'row' => 10],
                        ['action' => 'add', 'row' => 11],
                        ['action' => 'add', 'row' => 12],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A4'],
                        ['from' => 'A1', 'to' => 'A7'],
                        ['from' => 'A1', 'to' => 'A10'],
                        ['from' => 'B1', 'to' => 'B4'],
                        ['from' => 'B1', 'to' => 'B7'],
                        ['from' => 'B1', 'to' => 'B10'],
                        ['from' => 'C1', 'to' => 'C4'],
                        ['from' => 'C1', 'to' => 'C7'],
                        ['from' => 'C1', 'to' => 'C10'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C8'],
                        ['from' => 'C2', 'to' => 'C11'],
                    ],

                    'merge_cells' => ['A4:A6', 'B4:B5', 'A7:A9', 'B7:B8', 'A10:A12', 'B10:B11'],

                    'copy_width' => [],

                    'copy_cell_format' => [
                        ['from' => 'A1', 'to' => 'A4'],
                        ['from' => 'A1', 'to' => 'A7'],
                        ['from' => 'A1', 'to' => 'A10'],
                        ['from' => 'B1', 'to' => 'B4'],
                        ['from' => 'B1', 'to' => 'B7'],
                        ['from' => 'B1', 'to' => 'B10'],
                        ['from' => 'C1', 'to' => 'C4'],
                        ['from' => 'C1', 'to' => 'C7'],
                        ['from' => 'C1', 'to' => 'C10'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C8'],
                        ['from' => 'C2', 'to' => 'C11'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_multi_merge2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'B' => '[managers.name]',
                    ],
                    2 => [
                        'A' => '[managers.sales.month]',
                        'B' => '[managers.sales.amount]',
                    ],
                    3 => [
                        'B' => '[managers.sales.qty]',
                    ],
                    4 => [
                        'B' => 'Hello',
                    ],
                ],

                'data' => [
                    'managers' => [
                        'name' => [['Liam', 'Noah', 'Emma']],
                        'sales' => [
                            ['month' => '01', 'qty' => [[1, 2, 3]], 'amount' => [100, 101, 102]],
                            ['month' => '02', 'qty' => [[1, 2, 3]], 'amount' => [200, 201, 202]],
                            ['month' => '03', 'qty' => [[1, 2, 3]], 'amount' => [300, 301, 302]],
                        ],
                    ],
                ],

                'merge_cells' => ['B1:C1', 'A2:A4', 'B2:C2', 'B3:C3', 'B4:C4'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'B' => 'Liam',
                            'D' => 'Noah',
                            'F' => 'Emma',
                        ],
                        2 => [
                            'A' => '01',
                            'B' => 100,
                            'D' => 101,
                            'F' => 102,
                        ],
                        3 => [
                            'B' => 1,
                            'D' => 2,
                            'F' => 3,
                        ],
                        4 => [
                            'B' => 'Hello',
                        ],
                        5 => [
                            'A' => '02',
                            'B' => 200,
                            'D' => 201,
                            'F' => 202,
                        ],
                        6 => [
                            'B' => 1,
                            'D' => 2,
                            'F' => 3,
                        ],
                        8 => [
                            'A' => '03',
                            'B' => 300,
                            'D' => 301,
                            'F' => 302,
                        ],
                        9 => [
                            'B' => 1,
                            'D' => 2,
                            'F' => 3,
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 5],
                        ['action' => 'add', 'row' => 6],
                        ['action' => 'add', 'row' => 7],
                        ['action' => 'add', 'row' => 8],
                        ['action' => 'add', 'row' => 9],
                        ['action' => 'add', 'row' => 10],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'B1', 'to' => 'D1'],
                        ['from' => 'B1', 'to' => 'F1'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'F2'],
                        ['from' => 'B3', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B9'],
                        ['from' => 'B3', 'to' => 'D3'],
                        ['from' => 'B3', 'to' => 'F3'],
                        ['from' => 'C1', 'to' => 'E1'],
                        ['from' => 'C1', 'to' => 'G1'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C8'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'C2', 'to' => 'G2'],
                        ['from' => 'C3', 'to' => 'C6'],
                        ['from' => 'C3', 'to' => 'C9'],
                        ['from' => 'C3', 'to' => 'E3'],
                        ['from' => 'C3', 'to' => 'G3'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'D2', 'to' => 'D8'],
                        ['from' => 'D3', 'to' => 'D6'],
                        ['from' => 'D3', 'to' => 'D9'],
                        ['from' => 'E2', 'to' => 'E5'],
                        ['from' => 'E2', 'to' => 'E8'],
                        ['from' => 'E3', 'to' => 'E6'],
                        ['from' => 'E3', 'to' => 'E9'],
                        ['from' => 'F2', 'to' => 'F5'],
                        ['from' => 'F2', 'to' => 'F8'],
                        ['from' => 'F3', 'to' => 'F6'],
                        ['from' => 'F3', 'to' => 'F9'],
                        ['from' => 'G2', 'to' => 'G5'],
                        ['from' => 'G2', 'to' => 'G8'],
                        ['from' => 'G3', 'to' => 'G6'],
                        ['from' => 'G3', 'to' => 'G9'],
                    ],

                    'merge_cells' => [
                        'D1:E1', 'F1:G1',
                        'D2:E2', 'F2:G2',
                        'A5:A7', 'B5:C5', 'D5:E5', 'F5:G5',
                        'A8:A10', 'B8:C8', 'D8:E8', 'F8:G8',
                        'D3:E3', 'F3:G3',
                        'B6:C6', 'D6:E6', 'F6:G6',
                        'B9:C9', 'D9:E9', 'F9:G9',
                    ],

                    'copy_width' => [
                        ['from' => 'B', 'to' => 'D'],
                        ['from' => 'B', 'to' => 'F'],
                        ['from' => 'C', 'to' => 'E'],
                        ['from' => 'C', 'to' => 'G'],
                    ],

                    'copy_cell_format' => [
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'B1', 'to' => 'D1'],
                        ['from' => 'B1', 'to' => 'F1'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'F2'],
                        ['from' => 'B3', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B9'],
                        ['from' => 'B3', 'to' => 'D3'],
                        ['from' => 'B3', 'to' => 'F3'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'D2', 'to' => 'D8'],
                        ['from' => 'D3', 'to' => 'D6'],
                        ['from' => 'D3', 'to' => 'D9'],
                        ['from' => 'F2', 'to' => 'F5'],
                        ['from' => 'F2', 'to' => 'F8'],
                        ['from' => 'F3', 'to' => 'F6'],
                        ['from' => 'F3', 'to' => 'F9'],
                    ],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }
}
