<?php

namespace AnourValar\Office\Tests;

class TemplateParserTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @var \AnourValar\Office\Template\Parser
     */
    protected \AnourValar\Office\Template\Parser $service;

    /**
     * @see \PHPUnit\Framework\TestCase
     *
     * @return void
     */
    protected function setUp(): void
    {
        $this->service = new \AnourValar\Office\Template\Parser();
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
                    ],
                    2 => [
                        'A' => '[bar]',
                        'B' => 'test [bar] test',
                        'C' => 'bar',
                        'D' => '[a.b] -> [a.c] -> [a.d] -> [a.e.f]',
                        'H' => '[hello]',
                        'J' => '[world]',
                    ],
                    '3' => [
                        'D' => '[bar] [!bar]',
                    ],
                    '4' => [
                        'D' => 'hello world',
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
                        ],
                        2 => [
                            'A' => 'world',
                            'B' => 'test world test',
                            'D' => '11 -> 22 ->  -> oops',
                            'H' => null,
                            'J' => null,
                        ],
                    ],

                    'rows' => [['action' => 'delete', 'row' => 3]],

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
                "$id"
            );
        }

        $this->expectException(\LogicException::class);
        $this->service->schema([1 => ['A' => 'hello [world]']], ['world' => function () {}]);
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
                        'B' => '[bar] 1',
                    ],
                    4 => [
                        'A' => 'foo [= list.0]',
                        'B' => '[foo]',
                    ],
                    5 => [
                        'A' => 'bar',
                        'B' => '[bar] 2 [=bar]',
                    ],
                    6 => [
                        'A' => 'foo [= list.0.c]',
                        'B' => '[foo] [!list]',
                    ],
                    7 => [
                        'A' => 'bar',
                        'B' => '[bar] 3 [! list]',
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
                            'B' => 'test2 1',
                        ],
                        2 => [
                            'B' => 'test2 2',
                        ],
                        3 => [
                            'B' => 'test2 3',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 1],
                        ['action' => 'delete', 'row' => 1],
                        ['action' => 'delete', 'row' => 2],
                        ['action' => 'delete', 'row' => 3],
                    ],

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
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
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
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
                        ],
                        '2' => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 1],
                    ],

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.c]',
                        'D' => '[list.d]',
                        'E' => '[foo] hello [list.c] -> [list.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
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
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c]',
                        'D' => '[list_d]',
                        'E' => '[foo] hello [list_c] -> [list_d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
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
                        ],
                        2 => [
                            'B' => 'test1',
                            'C' => '11',
                            'D' => '12',
                            'E' => 'test1 hello 11 -> 12 world test2',
                            'G' => 'test2',
                        ],
                        3 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [['action' => 'delete', 'row' => 4]],

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list.c]',
                        'D' => '[list.d]',
                        'E' => '[foo] hello [list.c] -> [list.d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
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
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[list_c]',
                        'D' => '[list_d]',
                        'E' => '[foo] hello [list_c] -> [list_d] world [bar]',

                        'F' => 'bar',
                        'G' => '[bar]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
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
                        ],
                        2 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '11',
                            'D' => '12',
                            'E' => 'test1 hello 11 -> 12 world test2',
                            'F' => 'bar',
                            'G' => 'test2',
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '21',
                            'D' => '22',
                            'E' => 'test1 hello 21 -> 22 world test2',
                            'F' => 'bar',
                            'G' => 'test2',
                        ],
                        4 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                    ],

                    'copy_styles' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'B2', 'to' => 'B3'],
                        ['from' => 'C2', 'to' => 'C3'],
                        ['from' => 'D2', 'to' => 'D3'],
                        ['from' => 'E2', 'to' => 'E3'],
                        ['from' => 'F2', 'to' => 'F3'],
                        ['from' => 'G2', 'to' => 'G3'],
                    ],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => 11,
                            'D' => 12.5,
                            'E' => 'test1 hello 11 -> 12.5 world test2',
                            'F' => 'bar',
                            'G' => 'test2',
                        ],
                        3 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '21',
                            'D' => '22',
                            'E' => 'test1 hello 21 -> 22 world test2',
                            'F' => 'bar',
                            'G' => 'test2',
                        ],
                        4 => [
                            'A' => 'foo',
                            'B' => 'test1',
                            'C' => '31',
                            'D' => '32',
                            'E' => 'test1 hello 31 -> 32 world test2',
                            'F' => 'bar',
                            'G' => 'test2',
                        ],
                        5 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3],
                        ['action' => 'add', 'row' => 4],
                    ],

                    'copy_styles' => [
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
                    ],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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
                        ],
                        2 => [
                            'B' => 'test2',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 2],
                        ['action' => 'delete', 'row' => 2],
                    ],

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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

                    'copy_styles' => [],

                    'width' => [],
                ],
                $this->service->schema($item['values'], $item['data']),
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

            [
                'values' => [
                    1 => [
                        'A' => 'foo',
                        'B' => '[foo]',
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

                    'copy_styles' => [
                        ['from' => 'C2', 'to' => 'D2'],
                    ],

                    'width' => ['D' => 'C'],
                ],
                $this->service->schema($item['values'], $item['data']),
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

                    'copy_styles' => [
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                    ],

                    'width' => ['D' => 'C', 'E' => 'C'],
                ],
                $this->service->schema($item['values'], $item['data']),
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

                    'copy_styles' => [
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

                    'width' => ['D' => 'C', 'E' => 'C'],
                ],
                $this->service->schema($item['values'], $item['data']),
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

                    'copy_styles' => [
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

                    'width' => ['D' => 'C', 'E' => 'C'],
                ],
                $this->service->schema($item['values'], $item['data']),
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

                    'copy_styles' => [
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

                    'width' => ['F' => 'E', 'G' => 'E'],
                ],
                $this->service->schema($item['values'], $item['data']),
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
                        'E' => '[columns.0.one.0]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.0.b]',
                        'C' => '[rows.0.c]',
                        'E' => '[columns.0.two.0]',
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

                    'copy_styles' => [
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

                    'width' => ['F' => 'E', 'G' => 'E'],
                ],
                $this->service->schema($item['values'], $item['data']),
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
                        'B' => '[title]',
                    ],
                    3 => [
                        'G' => '[column.0.month.0] [= column.0.month.0]',
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
                        ['name' => 'Product 2', 'count' => 1, 'price' => 123]
                    ],
                    'column' => [
                        [
                            'month' => ['01', '02', '03'],
                            'amount' => [15000, 20000, 30000],
                        ],
                    ],
                ],
            ],

            [
                'values' => [
                    1 => [
                        'B' => '[title] [=column.month]',
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
                        ['name' => 'Product 2', 'count' => 1, 'price' => 123]
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
                            'C' => 'kg',
                            'D' => 753.14,
                            'E' => 'bar',
                            'G' => 15000,
                            'H' => 20000,
                            'I' => 30000,
                        ],
                        5 => [
                            'A' => 'Product 2',
                            'B' => 1,
                            'C' => 'kg',
                            'D' => 123,
                            'E' => 'bar',
                            'G' => null,
                            'H' => null,
                            'I' => null,
                        ],
                        7 => [
                            'B' => 3,
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 5],
                    ],

                    'copy_styles' => [
                        ['from' => 'A4', 'to' => 'A5'],
                        ['from' => 'B4', 'to' => 'B5'],
                        ['from' => 'C4', 'to' => 'C5'],
                        ['from' => 'D4', 'to' => 'D5'],
                        ['from' => 'E4', 'to' => 'E5'],
                        ['from' => 'G3', 'to' => 'H3'],
                        ['from' => 'G3', 'to' => 'I3'],
                        ['from' => 'G4', 'to' => 'G5'],
                        ['from' => 'G4', 'to' => 'H4'],
                        ['from' => 'G4', 'to' => 'I4'],
                        ['from' => 'H4', 'to' => 'H5'],
                        ['from' => 'I4', 'to' => 'I5'],
                    ],

                    'width' => ['H' => 'G', 'I' => 'G'],
                ],
                $this->service->schema($item['values'], $item['data']),
                "$id"
            );
        }
    }
}
