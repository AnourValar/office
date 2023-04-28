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
    public function test_collision_names()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[title]',
                    ],
                    2 => [
                        'A' => '[request.title]',
                    ],
                    3 => [
                        'A' => '[response.body]',
                    ],
                    4 => [
                        'A' => '[body]',
                    ],
                    5 => [
                        'A' => '[products.title.title]',
                    ],
                ],

                'data' => [
                    'title' => 'foo',
                    'request' => [],
                    'response' => ['body' => 'bar'],
                    'body' => [],
                    'products' => [
                        'title' => [111],
                    ],
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => '[title]',
                    ],
                    2 => [
                        'A' => '[request.title]',
                    ],
                    3 => [
                        'A' => '[response.body]',
                    ],
                    4 => [
                        'A' => '[body]',
                    ],
                    5 => [
                        'A' => '[products.title.title]',
                    ],
                ],

                'data' => [
                    'title' => 'foo',
                    'request' => null,
                    'response' => ['body' => 'bar'],
                    'body' => null,
                    'products' => [
                        'title' => [111],
                    ],
                ],
            ],

            [
                'values' => [
                    1 => [
                        'A' => '[title]',
                    ],
                    2 => [
                        'A' => '[request.title]',
                    ],
                    3 => [
                        'A' => '[response.body]',
                    ],
                    4 => [
                        'A' => '[body]',
                    ],
                    5 => [
                        'A' => '[products.title.title]',
                    ],
                ],

                'data' => [
                    'title' => 'foo',
                    'response' => ['body' => 'bar'],
                    'products' => [
                        'title' => [111],
                    ],
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'foo'],
                        2 => ['A' => null],
                        3 => ['A' => 'bar'],
                        4 => ['A' => null],
                        5 => ['A' => null],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
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
                        'Y' => null,
                    ],
                    '3' => [
                        'D' => '[bar] [!bar]',
                        'Y' => null,
                    ],
                    '4' => [
                        'D' => 'hello world',
                        'Y' => null,
                    ],
                    '5' => [
                        'A' => 1,
                        'B' => 2,
                        'Y' => null,
                    ],
                    '6' => [
                        'Q' => '=A1+B2+C3+D4+E5',
                        'Y' => null,
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

                    'rows' => [['action' => 'delete', 'row' => 3, 'qty' => 1]],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_not_scalar()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo]',
                        'B' => '[baz]',
                    ],
                    2 => [
                        'A' => 'hello [bar] world',
                        'B' => null,
                    ],
                    3 => [
                        'A' => '[test] [=foo]',
                        'B' => null,
                    ],
                    4 => [
                        'A' => '[test2] [=bar]',
                        'B' => null,
                    ],
                    5 => [
                        'A' => '[test2] [!foo]',
                        'B' => null,
                    ],
                ],

                'data' => [
                    'foo' => function () {},
                    'baz' => new \DateTime('2022-11-16'),
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
                            'B' => $item['data']['baz'],
                        ],
                        2 => [
                            'A' => 'hello  world',
                        ],
                        3 => [
                            'A' => $item['data']['test'],
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 4, 'qty' => 1],
                        ['action' => 'delete', 'row' => 4, 'qty' => 1],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
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
            ],
            $this->service->schema([1 => ['A' => 'hello [world]']], ['world' => function () {}], [])->toArray()
        );
    }

   /**
     * @return void
     */
    public function test_schema_conditions1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'AA [$= foo] [$= bar]',
                        'B' => 'BB [$= baz] [$= foo] [$= bar]',
                        'C' => 'CC [$= foo] [$= baz] [$= bar]',
                        'D' => 'DD [$= foo] [$= bar] [$= baz]',
                    ],

                    2 => [
                        'A' => 'AA [$! baz] [$! foobar]',
                        'B' => 'BB [$! foo] [$! baz] [$! foobar]',
                        'C' => 'CC [$! baz] [$! foo] [$! foobar]',
                        'D' => 'DD [$! baz] [$! foobar] [$! foo]',
                    ],
                ],

                'data' => [
                    'foo' => 'foo',
                    'bar' => 'bar',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => 'AA',
                            'B' => null,
                            'C' => null,
                            'D' => null,
                        ],
                        2 => [
                            'A' => 'AA',
                            'B' => null,
                            'C' => null,
                            'D' => null,
                        ],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
    }

   /**
     * @return void
     */
    public function test_schema_conditions2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '11 [= foo] [= bar]',
                    ],
                    2 => [
                        'A' => '22 [= baz] [= foo] [= bar]',
                    ],
                    3 => [
                        'A' => '33 [= foo] [= baz] [= bar]',
                    ],
                    4 => [
                        'A' => '44 [= foo] [= bar] [= baz]',
                    ],

                    5 => [
                        'A' => '55 [! baz] [! foobar]',
                    ],
                    6 => [
                        'A' => '66 [! foo] [! baz] [! foobar]',
                    ],
                    7 => [
                        'A' => '77 [! baz] [! foo] [! foobar]',
                    ],
                    8 => [
                        'A' => '88 [! baz] [! foobar] [! foo]',
                    ],
                ],

                'data' => [
                    'foo' => 'foo',
                    'bar' => 'bar',
                ],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => [
                            'A' => '11',
                        ],
                        2 => [
                            'A' => '55',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],

                        ['action' => 'delete', 'row' => 3, 'qty' => 1],
                        ['action' => 'delete', 'row' => 3, 'qty' => 1],
                        ['action' => 'delete', 'row' => 3, 'qty' => 1],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], [])->toArray(),
                "$id"
            );
        }
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo] [= list]',
                        'C' => null,
                    ],
                    3 => [
                        'A' => 'bar [= bar]',
                        'B' => '[bar] 111',
                        'C' => null,
                    ],
                    4 => [
                        'A' => 'foo [= list.0]',
                        'B' => '[foo]',
                        'C' => null,
                    ],
                    5 => [
                        'A' => 'bar',
                        'B' => '[bar] 222 [=bar]',
                        'C' => '=A3+B5+C7',
                    ],
                    6 => [
                        'A' => 'foo [= list.0.c]',
                        'B' => '[foo] [!list]',
                        'C' => null,
                    ],
                    7 => [
                        'A' => 'bar',
                        'B' => '[bar] 333 [! list]',
                        'C' => null,
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
                        ['action' => 'delete', 'row' => 1, 'qty' => 1],
                        ['action' => 'delete', 'row' => 1, 'qty' => 1],
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],
                        ['action' => 'delete', 'row' => 3, 'qty' => 1],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        ['action' => 'delete', 'row' => 1, 'qty' => 1],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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

                    'rows' => [['action' => 'delete', 'row' => 4, 'qty' => 1]],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 1],
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
                        ['from' => 'I2', 'to' => 'I3'],
                        ['from' => 'J2', 'to' => 'J3'],
                        ['from' => 'K2', 'to' => 'K3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A4'],
                        ['from' => 'B2', 'to' => 'B3:B4'],
                        ['from' => 'C2', 'to' => 'C3:C4'],
                        ['from' => 'D2', 'to' => 'D3:D4'],
                        ['from' => 'E2', 'to' => 'E3:E4'],
                        ['from' => 'G2', 'to' => 'G3:G4'],
                        ['from' => 'H2', 'to' => 'H3:H4'],
                        ['from' => 'I2', 'to' => 'I3:I4'],
                        ['from' => 'J2', 'to' => 'J3:J4'],
                        ['from' => 'K2', 'to' => 'K3:K4'],
                    ],

                    'merge_cells' => ['E3:F3', 'E4:F4'],

                    'copy_width' => [],
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
                        'G' => null,
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
                        'G' => null,
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
                        'G' => null,
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
                        'G' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        'K' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 3],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A5'],
                        ['from' => 'B2', 'to' => 'B3:B5'],
                        ['from' => 'C2', 'to' => 'C3:C5'],
                        ['from' => 'D2', 'to' => 'D3:D5'],
                        ['from' => 'E2', 'to' => 'E3:E5'],
                        ['from' => 'G2', 'to' => 'G3:G5'],
                        ['from' => 'H2', 'to' => 'H3:H5'],
                        ['from' => 'I2', 'to' => 'I3:I5'],
                        ['from' => 'J2', 'to' => 'J3:J5'],
                        ['from' => 'K2', 'to' => 'K3:K5'],
                    ],

                    'merge_cells' => ['E3:F3', 'E4:F4', 'E5:F5'],

                    'copy_width' => [],
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
                        'C' => null,
                    ],
                    4 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => null,
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
                        'C' => null,
                    ],
                    4 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => null,
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
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],
                    ],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.c.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => null,
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[foo]',

                        'C' => '[matrix.0.0.0]',
                    ],
                    3 => [
                        'A' => 'bar',
                        'B' => '[bar]',
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'C' => null,
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
                        'D' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.b]',
                        'C' => '[matrix.c]',
                        'D' => null,
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
                        'D' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.*.b]',
                        'C' => '[matrix.*.c]',
                        'D' => null,
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
                        'D' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.*.b]',
                        'C' => '[matrix.*.c]',
                        'D' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A4'],
                        ['from' => 'B2', 'to' => 'B3:B4'],
                        ['from' => 'C2', 'to' => 'C3:C4'],
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'D2', 'to' => 'D3:D4'],
                        ['from' => 'E2', 'to' => 'E3:E4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'C', 'to' => 'D'],
                        ['from' => 'C', 'to' => 'E'],
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
                        'D' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.0.b]',
                        'C' => '[matrix.0.c.0]',
                        'D' => null,
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
                        'D' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.0.b] [= matrix.0.b]',
                        'C' => '[matrix.0.c.0] [=matrix.c]',
                        'D' => null,
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
                        'D' => null,
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[matrix.b]',
                        'C' => '[matrix.c]',
                        'D' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A4'],
                        ['from' => 'B2', 'to' => 'B3:B4'],
                        ['from' => 'C2', 'to' => 'C3:C4'],
                        ['from' => 'C2', 'to' => 'D2'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'D2', 'to' => 'D3:D4'],
                        ['from' => 'E2', 'to' => 'E3:E4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'C', 'to' => 'D'],
                        ['from' => 'C', 'to' => 'E'],
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
                        'E' => null,
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
                        'E' => null,
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
                        'E' => null,
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
                        'E' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A4'],
                        ['from' => 'B2', 'to' => 'B3:B4'],
                        ['from' => 'C2', 'to' => 'C3:C4'],
                        ['from' => 'D2', 'to' => 'D3:D4'],
                        ['from' => 'E2', 'to' => 'E3:E4'],
                        ['from' => 'E2', 'to' => 'F2'],
                        ['from' => 'E2', 'to' => 'G2'],
                        ['from' => 'F2', 'to' => 'F3:F4'],
                        ['from' => 'G2', 'to' => 'G3:G4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'E', 'to' => 'F'],
                        ['from' => 'E', 'to' => 'G'],
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
                        'E' => null,
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
                        'E' => null,
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
                        'E' => null,
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
                        'E' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 1],
                    ],

                    'copy_style' => [
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

                    'merge_cells' => ['C3:D3'],

                    'copy_width' => [
                        ['from' => 'E', 'to' => 'F'],
                        ['from' => 'E', 'to' => 'G'],
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
                        'E' => '[columns.*.one.*]',
                    ],
                    2 => [
                        'A' => 'foo',
                        'B' => '[rows.1.b]',
                        'C' => '[rows.1.c]',
                        'E' => '[columns.*.two.*]',
                    ],
                    3 => [
                        'A' => '[foo]',
                        'E' => null,
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
                        'Q' => null,
                    ],
                    3 => [
                        'G' => '[column.month] [= column.month]',
                        'Q' => null,
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
                        'Q' => null,
                    ],
                    7 => [
                        'C' => '=D6 [oops]',
                        'D' => 'D6',
                        'Q' => null,
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
                        'Q' => null,
                    ],
                    3 => [
                        'G' => '[column.*.month] [= column.*.month.*]',
                        'Q' => null,
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
                        'Q' => null,
                    ],
                    7 => [
                        'C' => '=D6 [oops]',
                        'D' => 'D6',
                        'Q' => null,
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
                        ['action' => 'add', 'row' => 5, 'qty' => 1],
                    ],

                    'copy_style' => [
                        ['from' => 'A4', 'to' => 'A5'],
                        ['from' => 'B4', 'to' => 'B5'],
                        ['from' => 'C4', 'to' => 'C5'],
                        ['from' => 'D4', 'to' => 'D5'],
                        ['from' => 'E4', 'to' => 'E5'],
                        ['from' => 'F4', 'to' => 'F5'],
                        ['from' => 'G3', 'to' => 'H3'],
                        ['from' => 'G3', 'to' => 'I3'],
                        ['from' => 'G4', 'to' => 'G5'],
                        ['from' => 'G4', 'to' => 'I4'],
                        ['from' => 'G4', 'to' => 'K4'],
                        ['from' => 'I4', 'to' => 'I5'],
                        ['from' => 'K4', 'to' => 'K5'],
                        ['from' => 'M4', 'to' => 'M5'],
                        ['from' => 'N4', 'to' => 'N5'],
                        ['from' => 'O4', 'to' => 'O5'],
                        ['from' => 'P4', 'to' => 'P5'],
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
                        'G' => null,
                    ],
                    3 => [
                        'G' => '[column.*.month.*] [= column.0.month]',
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
                        'G' => null,
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
                        'F' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 3],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A5'],
                        ['from' => 'B2', 'to' => 'B3:B5'],
                        ['from' => 'C2', 'to' => 'C3:C5'],
                        ['from' => 'D2', 'to' => 'D3:D5'],
                        ['from' => 'E2', 'to' => 'E3:E5'],
                        ['from' => 'F2', 'to' => 'F3:F5'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'K' => null,
                    ],
                    3 => [
                        'A' => '[table.price]',
                        'B' => '[table.count]',
                        'C' => '=A3*B3',
                        'K' => null,
                    ],
                    4 => [
                        'A' => '[=hello]',
                        'K' => null,
                    ],
                    5 => [
                        'A' => '[!table]',
                        'K' => null,
                    ],
                    6 => [
                        'A' => '[!table.count]',
                        'K' => null,
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
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],
                        ['action' => 'add', 'row' => 3, 'qty' => 4],
                        ['action' => 'delete', 'row' => 7, 'qty' => 1],
                        ['action' => 'delete', 'row' => 7, 'qty' => 1],
                        ['action' => 'delete', 'row' => 7, 'qty' => 1],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A6'],
                        ['from' => 'B2', 'to' => 'B3:B6'],
                        ['from' => 'C2', 'to' => 'C3:C6'],
                        ['from' => 'D2', 'to' => 'D3:D6'],
                        ['from' => 'E2', 'to' => 'E3:E6'],
                        ['from' => 'F2', 'to' => 'F3:F6'],
                        ['from' => 'G2', 'to' => 'G3:G6'],
                        ['from' => 'H2', 'to' => 'H3:H6'],
                        ['from' => 'I2', 'to' => 'I3:I6'],
                        ['from' => 'J2', 'to' => 'J3:J6'],
                        ['from' => 'K2', 'to' => 'K3:K6'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'G' => null,
                    ],
                    3 => [
                        'A' => '[table.price]',
                        'B' => '[table.count]',
                        'C' => '=A3*B3+C1+D2+E4+F5',
                        'G' => null,
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
                        ['action' => 'delete', 'row' => 2, 'qty' => 1],
                        ['action' => 'add', 'row' => 3, 'qty' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A6'],
                        ['from' => 'B2', 'to' => 'B3:B6'],
                        ['from' => 'C2', 'to' => 'C3:C6'],
                        ['from' => 'D2', 'to' => 'D3:D6'],
                        ['from' => 'E2', 'to' => 'E3:E6'],
                        ['from' => 'F2', 'to' => 'F3:F6'],
                        ['from' => 'G2', 'to' => 'G3:G6'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A4'],
                        ['from' => 'A2', 'to' => 'B2'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'B2', 'to' => 'B3:B4'],
                        ['from' => 'C2', 'to' => 'C3:C4'],
                        ['from' => 'D2', 'to' => 'D3:D4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [
                        ['from' => 'A', 'to' => 'B'],
                        ['from' => 'A', 'to' => 'C'],
                        ['from' => 'A', 'to' => 'D'],
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
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A4'],
                        ['from' => 'A2', 'to' => 'C2'],
                        ['from' => 'A2', 'to' => 'E2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'C2', 'to' => 'C3:C4'],
                        ['from' => 'E2', 'to' => 'E3:E4'],
                        ['from' => 'G2', 'to' => 'G3:G4'],
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
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3:A4'],
                        ['from' => 'A2', 'to' => 'D2'],
                        ['from' => 'A2', 'to' => 'G2'],
                        ['from' => 'A2', 'to' => 'J2'],
                        ['from' => 'D2', 'to' => 'D3:D4'],
                        ['from' => 'G2', 'to' => 'G3:G4'],
                        ['from' => 'J2', 'to' => 'J3:J4'],
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
                        ['action' => 'add', 'row' => 4, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A3', 'to' => 'A4:A5'],
                        ['from' => 'C2', 'to' => 'E2'],
                        ['from' => 'C2', 'to' => 'G2'],
                        ['from' => 'C3', 'to' => 'C4:C5'],
                        ['from' => 'C3', 'to' => 'E3'],
                        ['from' => 'C3', 'to' => 'G3'],
                        ['from' => 'E3', 'to' => 'E4:E5'],
                        ['from' => 'G3', 'to' => 'G4:G5'],
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
                        ['action' => 'add', 'row' => 4, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A3', 'to' => 'A4:A5'],
                        ['from' => 'D2', 'to' => 'G2'],
                        ['from' => 'D2', 'to' => 'J2'],
                        ['from' => 'D3', 'to' => 'D4:D5'],
                        ['from' => 'D3', 'to' => 'G3'],
                        ['from' => 'D3', 'to' => 'J3'],
                        ['from' => 'G3', 'to' => 'G4:G5'],
                        ['from' => 'J3', 'to' => 'J4:J5'],
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                        'C' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                        'C' => null,
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                        'C' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                        'C' => null,
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                        'C' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                        'C' => null,
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
                        ['action' => 'add', 'row' => 4, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C3', 'to' => 'C5'],
                    ],

                    'merge_cells' => [
                        'A4:A5',
                    ],

                    'copy_width' => [],
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
                        'B' => null,
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
                        ['action' => 'add', 'row' => 3, 'qty' => 1],
                    ],

                    'copy_style' => [
                        ['from' => 'B2', 'to' => 'B3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                        'C' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                        'C' => null,
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
                        ['action' => 'add', 'row' => 4, 'qty' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A4'],
                        ['from' => 'A2', 'to' => 'A6'],
                        ['from' => 'B2', 'to' => 'B4'],
                        ['from' => 'B2', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B5'],
                        ['from' => 'B3', 'to' => 'B7'],
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C6'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'C7'],
                    ],

                    'merge_cells' => [
                        'A4:A5', 'A6:A7',
                    ],

                    'copy_width' => [],
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
                        'C' => null,
                    ],
                    2 => [
                        'A' => '[project.id]',
                        'B' => '[project.name]',
                        'C' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => '[project.amount_2]',
                    ],
                    4 => [
                        'A' => '[foo]',
                        'C' => null,
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
                        ['action' => 'add', 'row' => 4, 'qty' => 6],
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
                        ['from' => 'C2', 'to' => 'C4'],
                        ['from' => 'C2', 'to' => 'C6'],
                        ['from' => 'C2', 'to' => 'C8'],
                        ['from' => 'C3', 'to' => 'C5'],
                        ['from' => 'C3', 'to' => 'C7'],
                        ['from' => 'C3', 'to' => 'C9'],
                    ],

                    'merge_cells' => [
                        'A4:A5', 'A6:A7', 'A8:A9',
                    ],

                    'copy_width' => [],
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
                        'C' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'C' => null,
                    ],
                    4 => [
                        'C' => '[project.amount_2]',
                    ],
                    6 => [
                        'A' => '[foo]',
                        'C' => null,
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
                        ['action' => 'add', 'row' => 5, 'qty' => 9],
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
                        ['from' => 'B4', 'to' => 'B7'],
                        ['from' => 'B4', 'to' => 'B10'],
                        ['from' => 'B4', 'to' => 'B13'],
                        ['from' => 'C2', 'to' => 'C5'],
                        ['from' => 'C2', 'to' => 'C8'],
                        ['from' => 'C2', 'to' => 'C11'],
                        ['from' => 'C3', 'to' => 'C6'],
                        ['from' => 'C3', 'to' => 'C9'],
                        ['from' => 'C3', 'to' => 'C12'],
                        ['from' => 'C4', 'to' => 'C7'],
                        ['from' => 'C4', 'to' => 'C10'],
                        ['from' => 'C4', 'to' => 'C13'],
                    ],

                    'merge_cells' => [
                        'A5:A7', 'A8:A10', 'A11:A13',
                    ],

                    'copy_width' => [],
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
                        'D' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'D' => null,
                    ],
                    4 => [
                        'C' => '[project.amount_2]',
                        'D' => null,
                    ],
                    6 => [
                        'A' => '[foo]',
                        'D' => null,
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
                        ['action' => 'add', 'row' => 6, 'qty' => 12],
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
                        ['from' => 'B4', 'to' => 'B8'],
                        ['from' => 'B4', 'to' => 'B12'],
                        ['from' => 'B4', 'to' => 'B16'],
                        ['from' => 'B5', 'to' => 'B9'],
                        ['from' => 'B5', 'to' => 'B13'],
                        ['from' => 'B5', 'to' => 'B17'],
                        ['from' => 'C2', 'to' => 'C6'],
                        ['from' => 'C2', 'to' => 'C10'],
                        ['from' => 'C2', 'to' => 'C14'],
                        ['from' => 'C3', 'to' => 'C7'],
                        ['from' => 'C3', 'to' => 'C11'],
                        ['from' => 'C3', 'to' => 'C15'],
                        ['from' => 'C4', 'to' => 'C8'],
                        ['from' => 'C4', 'to' => 'C12'],
                        ['from' => 'C4', 'to' => 'C16'],
                        ['from' => 'C5', 'to' => 'C9'],
                        ['from' => 'C5', 'to' => 'C13'],
                        ['from' => 'C5', 'to' => 'C17'],
                        ['from' => 'D2', 'to' => 'D6'],
                        ['from' => 'D2', 'to' => 'D10'],
                        ['from' => 'D2', 'to' => 'D14'],
                        ['from' => 'D3', 'to' => 'D7'],
                        ['from' => 'D3', 'to' => 'D11'],
                        ['from' => 'D3', 'to' => 'D15'],
                        ['from' => 'D4', 'to' => 'D8'],
                        ['from' => 'D4', 'to' => 'D12'],
                        ['from' => 'D4', 'to' => 'D16'],
                        ['from' => 'D5', 'to' => 'D9'],
                        ['from' => 'D5', 'to' => 'D13'],
                        ['from' => 'D5', 'to' => 'D17'],
                    ],

                    'merge_cells' => [
                        'A6:A9', 'A10:A13', 'A14:A17',
                    ],

                    'copy_width' => [],
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
                        'E' => null,
                    ],
                    3 => [
                        'B' => '[project.amount_1]',
                        'E' => null,
                    ],
                    5 => [
                        'C' => '[project.amount_2]',
                        'E' => null,
                    ],
                    7 => [
                        'A' => '[foo]',
                        'E' => null,
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
                        ['action' => 'add', 'row' => 7, 'qty' => 15],
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
                        ['from' => 'B4', 'to' => 'B9'],
                        ['from' => 'B4', 'to' => 'B14'],
                        ['from' => 'B4', 'to' => 'B19'],
                        ['from' => 'B5', 'to' => 'B10'],
                        ['from' => 'B5', 'to' => 'B15'],
                        ['from' => 'B5', 'to' => 'B20'],
                        ['from' => 'B6', 'to' => 'B11'],
                        ['from' => 'B6', 'to' => 'B16'],
                        ['from' => 'B6', 'to' => 'B21'],
                        ['from' => 'C2', 'to' => 'C7'],
                        ['from' => 'C2', 'to' => 'C12'],
                        ['from' => 'C2', 'to' => 'C17'],
                        ['from' => 'C3', 'to' => 'C8'],
                        ['from' => 'C3', 'to' => 'C13'],
                        ['from' => 'C3', 'to' => 'C18'],
                        ['from' => 'C4', 'to' => 'C9'],
                        ['from' => 'C4', 'to' => 'C14'],
                        ['from' => 'C4', 'to' => 'C19'],
                        ['from' => 'C5', 'to' => 'C10'],
                        ['from' => 'C5', 'to' => 'C15'],
                        ['from' => 'C5', 'to' => 'C20'],
                        ['from' => 'C6', 'to' => 'C11'],
                        ['from' => 'C6', 'to' => 'C16'],
                        ['from' => 'C6', 'to' => 'C21'],
                        ['from' => 'D2', 'to' => 'D7'],
                        ['from' => 'D2', 'to' => 'D12'],
                        ['from' => 'D2', 'to' => 'D17'],
                        ['from' => 'D3', 'to' => 'D8'],
                        ['from' => 'D3', 'to' => 'D13'],
                        ['from' => 'D3', 'to' => 'D18'],
                        ['from' => 'D4', 'to' => 'D9'],
                        ['from' => 'D4', 'to' => 'D14'],
                        ['from' => 'D4', 'to' => 'D19'],
                        ['from' => 'D5', 'to' => 'D10'],
                        ['from' => 'D5', 'to' => 'D15'],
                        ['from' => 'D5', 'to' => 'D20'],
                        ['from' => 'D6', 'to' => 'D11'],
                        ['from' => 'D6', 'to' => 'D16'],
                        ['from' => 'D6', 'to' => 'D21'],
                        ['from' => 'E2', 'to' => 'E7'],
                        ['from' => 'E2', 'to' => 'E12'],
                        ['from' => 'E2', 'to' => 'E17'],
                        ['from' => 'E3', 'to' => 'E8'],
                        ['from' => 'E3', 'to' => 'E13'],
                        ['from' => 'E3', 'to' => 'E18'],
                        ['from' => 'E4', 'to' => 'E9'],
                        ['from' => 'E4', 'to' => 'E14'],
                        ['from' => 'E4', 'to' => 'E19'],
                        ['from' => 'E5', 'to' => 'E10'],
                        ['from' => 'E5', 'to' => 'E15'],
                        ['from' => 'E5', 'to' => 'E20'],
                        ['from' => 'E6', 'to' => 'E11'],
                        ['from' => 'E6', 'to' => 'E16'],
                        ['from' => 'E6', 'to' => 'E21'],
                    ],

                    'merge_cells' => [
                        'A7:A11', 'A12:A16', 'A17:A21',
                    ],

                    'copy_width' => [],
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
                        'B' => null,
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
                        ['action' => 'add', 'row' => 5, 'qty' => 6],
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
                        'B' => null,
                    ],
                    2 => [
                        'A' => '=SUM(A1:A1)',
                        'B' => '=SUM(A1:A10)',
                    ],
                    3 => [
                        'A' => '-',
                        'B' => '[! list]',
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
                        ['action' => 'add', 'row' => 2, 'qty' => 1],
                        ['action' => 'delete', 'row' => 4, 'qty' => 1],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2'],
                        ['from' => 'B1', 'to' => 'B2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        'B' => null,
                    ],
                    3 => [
                        'A' => '=SUM(A2:A2)',
                        'B' => '=SUM(A2:A20)',
                    ],
                    4 => [
                        'A' => '-',
                        'B' => '[! list]',
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
                        ['action' => 'add', 'row' => 3, 'qty' => 1],
                        ['action' => 'delete', 'row' => 5, 'qty' => 1],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                        ['from' => 'B2', 'to' => 'B3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        ['action' => 'add', 'row' => 2, 'qty' => 1],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        ['action' => 'add', 'row' => 3, 'qty' => 1],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
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
                        ['action' => 'add', 'row' => 4, 'qty' => 9],
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
                        7 => [
                            'B' => 'Hello',
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
                        10 => [
                            'B' => 'Hello',
                        ],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 5, 'qty' => 6],
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
                        ['from' => 'B4', 'to' => 'B7'],
                        ['from' => 'B4', 'to' => 'B10'],
                        ['from' => 'D2', 'to' => 'D5'],
                        ['from' => 'D2', 'to' => 'D8'],
                        ['from' => 'D3', 'to' => 'D6'],
                        ['from' => 'D3', 'to' => 'D9'],
                        ['from' => 'F2', 'to' => 'F5'],
                        ['from' => 'F2', 'to' => 'F8'],
                        ['from' => 'F3', 'to' => 'F6'],
                        ['from' => 'F3', 'to' => 'F9'],
                    ],

                    'merge_cells' => [
                        'D1:E1', 'F1:G1',
                        'D2:E2', 'F2:G2',
                        'A5:A7', 'B5:C5', 'D5:E5', 'F5:G5',
                        'A8:A10', 'B8:C8', 'D8:E8', 'F8:G8',
                        'D3:E3', 'F3:G3',
                        'B6:C6', 'D6:E6', 'F6:G6',
                        'B9:C9', 'D9:E9', 'F9:G9',
                        'B7:C7', 'B10:C10',
                    ],

                    'copy_width' => [
                        ['from' => 'B', 'to' => 'D'],
                        ['from' => 'B', 'to' => 'F'],
                        ['from' => 'C', 'to' => 'E'],
                        ['from' => 'C', 'to' => 'G'],
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
    public function test_schema_list_long()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[foo1]',
                        'B' => '[bar1]',
                        'C' => 'foo1',
                        'D' => 'bar1',
                    ],
                    10 => [
                        'A' => '[foo2]',
                        'B' => '[bar2]',
                        'C' => 'foo2',
                        'D' => 'bar2',
                    ],
                ],

                'data' => [
                    'foo1' => [
                        'foo1-1', 'foo1-2', 'foo1-3', 'foo1-4', 'foo1-5', 'foo1-6', 'foo1-7', 'foo1-8',
                        'foo1-9', 'foo1-10', 'foo1-11', 'foo1-12', 'foo1-13', 'foo1-14', 'foo1-15', 'foo1-16',
                    ],
                    'bar1' => [
                        'bar1-1', 'bar1-2', 'bar1-3', 'bar1-4', 'bar1-5', 'bar1-6', 'bar1-7', 'bar1-8',
                        'bar1-9', 'bar1-10', 'bar1-11', 'bar1-12', 'bar1-13', 'bar1-14', 'bar1-15', 'bar1-16',
                    ],

                    'foo2' => [
                        'foo2-1', 'foo2-2', 'foo2-3', 'foo2-4', 'foo2-5', 'foo2-6', 'foo2-7', 'foo2-8',
                        'foo2-9', 'foo2-10', 'foo2-11', 'foo2-12', 'foo2-13', 'foo2-14', 'foo2-15', 'foo2-16',
                    ],
                    'bar2' => [
                        'bar2-1', 'bar2-2', 'bar2-3', 'bar2-4', 'bar2-5', 'bar2-6', 'bar2-7', 'bar2-8',
                        'bar2-9', 'bar2-10', 'bar2-11', 'bar2-12', 'bar2-13', 'bar2-14', 'bar2-15', 'bar2-16',
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'foo1-1', 'B' => 'bar1-1', 'C' => 'foo1', 'D' => 'bar1'],
                        2 => ['A' => 'foo1-2', 'B' => 'bar1-2', 'C' => 'foo1', 'D' => 'bar1'],
                        3 => ['A' => 'foo1-3', 'B' => 'bar1-3', 'C' => 'foo1', 'D' => 'bar1'],
                        4 => ['A' => 'foo1-4', 'B' => 'bar1-4', 'C' => 'foo1', 'D' => 'bar1'],
                        5 => ['A' => 'foo1-5', 'B' => 'bar1-5', 'C' => 'foo1', 'D' => 'bar1'],
                        6 => ['A' => 'foo1-6', 'B' => 'bar1-6', 'C' => 'foo1', 'D' => 'bar1'],
                        7 => ['A' => 'foo1-7', 'B' => 'bar1-7', 'C' => 'foo1', 'D' => 'bar1'],
                        8 => ['A' => 'foo1-8', 'B' => 'bar1-8', 'C' => 'foo1', 'D' => 'bar1'],
                        9 => ['A' => 'foo1-9', 'B' => 'bar1-9', 'C' => 'foo1', 'D' => 'bar1'],
                        10 => ['A' => 'foo1-10', 'B' => 'bar1-10', 'C' => 'foo1', 'D' => 'bar1'],
                        11 => ['A' => 'foo1-11', 'B' => 'bar1-11', 'C' => 'foo1', 'D' => 'bar1'],
                        12 => ['A' => 'foo1-12', 'B' => 'bar1-12', 'C' => 'foo1', 'D' => 'bar1'],
                        13 => ['A' => 'foo1-13', 'B' => 'bar1-13', 'C' => 'foo1', 'D' => 'bar1'],
                        14 => ['A' => 'foo1-14', 'B' => 'bar1-14', 'C' => 'foo1', 'D' => 'bar1'],
                        15 => ['A' => 'foo1-15', 'B' => 'bar1-15', 'C' => 'foo1', 'D' => 'bar1'],
                        16 => ['A' => 'foo1-16', 'B' => 'bar1-16', 'C' => 'foo1', 'D' => 'bar1'],

                        25 => ['A' => 'foo2-1', 'B' => 'bar2-1', 'C' => 'foo2', 'D' => 'bar2'],
                        26 => ['A' => 'foo2-2', 'B' => 'bar2-2', 'C' => 'foo2', 'D' => 'bar2'],
                        27 => ['A' => 'foo2-3', 'B' => 'bar2-3', 'C' => 'foo2', 'D' => 'bar2'],
                        28 => ['A' => 'foo2-4', 'B' => 'bar2-4', 'C' => 'foo2', 'D' => 'bar2'],
                        29 => ['A' => 'foo2-5', 'B' => 'bar2-5', 'C' => 'foo2', 'D' => 'bar2'],
                        30 => ['A' => 'foo2-6', 'B' => 'bar2-6', 'C' => 'foo2', 'D' => 'bar2'],
                        31 => ['A' => 'foo2-7', 'B' => 'bar2-7', 'C' => 'foo2', 'D' => 'bar2'],
                        32 => ['A' => 'foo2-8', 'B' => 'bar2-8', 'C' => 'foo2', 'D' => 'bar2'],
                        33 => ['A' => 'foo2-9', 'B' => 'bar2-9', 'C' => 'foo2', 'D' => 'bar2'],
                        34 => ['A' => 'foo2-10', 'B' => 'bar2-10', 'C' => 'foo2', 'D' => 'bar2'],
                        35 => ['A' => 'foo2-11', 'B' => 'bar2-11', 'C' => 'foo2', 'D' => 'bar2'],
                        36 => ['A' => 'foo2-12', 'B' => 'bar2-12', 'C' => 'foo2', 'D' => 'bar2'],
                        37 => ['A' => 'foo2-13', 'B' => 'bar2-13', 'C' => 'foo2', 'D' => 'bar2'],
                        38 => ['A' => 'foo2-14', 'B' => 'bar2-14', 'C' => 'foo2', 'D' => 'bar2'],
                        39 => ['A' => 'foo2-15', 'B' => 'bar2-15', 'C' => 'foo2', 'D' => 'bar2'],
                        40 => ['A' => 'foo2-16', 'B' => 'bar2-16', 'C' => 'foo2', 'D' => 'bar2'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 15],
                        ['action' => 'add', 'row' => 26, 'qty' => 15],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A16'],
                        ['from' => 'A25', 'to' => 'A26:A40'],
                        ['from' => 'B1', 'to' => 'B2:B16'],
                        ['from' => 'B25', 'to' => 'B26:B40'],
                        ['from' => 'C1', 'to' => 'C2:C16'],
                        ['from' => 'C25', 'to' => 'C26:C40'],
                        ['from' => 'D1', 'to' => 'D2:D16'],
                        ['from' => 'D25', 'to' => 'D26:D40'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_several_tables1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[tableOne.a]',
                        'C' => '',
                        'D' => '[tableOne.d]',
                    ],
                    3 => [
                        'A' => '[tableTwo.a]',
                        'C' => '',
                        'D' => '[tableTwo.d]',
                    ],
                    5 => [
                        'A' => '[tableThree.a]',
                        'C' => '',
                        'D' => '[tableThree.d]',
                    ],
                ],

                'data' => [
                    'tableOne' => [
                        'a' => ['one-a-1', 'one-a-2', 'one-a-3'],
                        'd' => ['one-d-1', 'one-d-2', 'one-d-3'],
                    ],
                    'tableTwo' => [
                        'a' => ['two-a-1', 'two-a-2', 'two-a-3'],
                        'd' => ['two-d-1', 'two-d-2', 'two-d-3'],
                    ],
                    'tableThree' => [
                        'a' => ['three-a-1', 'three-a-2', 'three-a-3'],
                        'd' => ['three-d-1', 'three-d-2', 'three-d-3'],
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'one-a-1', 'D' => 'one-d-1'],
                        2 => ['A' => 'one-a-2', 'D' => 'one-d-2'],
                        3 => ['A' => 'one-a-3', 'D' => 'one-d-3'],
                        5 => ['A' => 'two-a-1', 'D' => 'two-d-1'],
                        6 => ['A' => 'two-a-2', 'D' => 'two-d-2'],
                        7 => ['A' => 'two-a-3', 'D' => 'two-d-3'],
                        9 => ['A' => 'three-a-1', 'D' => 'three-d-1'],
                        10 => ['A' => 'three-a-2', 'D' => 'three-d-2'],
                        11 => ['A' => 'three-a-3', 'D' => 'three-d-3'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                        ['action' => 'add', 'row' => 6, 'qty' => 2],
                        ['action' => 'add', 'row' => 10, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'A5', 'to' => 'A6:A7'],
                        ['from' => 'A9', 'to' => 'A10:A11'],
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'B5', 'to' => 'B6:B7'],
                        ['from' => 'B9', 'to' => 'B10:B11'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'C5', 'to' => 'C6:C7'],
                        ['from' => 'C9', 'to' => 'C10:C11'],
                        ['from' => 'D1', 'to' => 'D2:D3'],
                        ['from' => 'D5', 'to' => 'D6:D7'],
                        ['from' => 'D9', 'to' => 'D10:D11'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_several_tables2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[tableOne.a]',
                        'C' => null,
                        'D' => '[tableOne.d]',
                    ],
                    3 => [
                        'A' => '[tableTwo.a]',
                        'C' => null,
                        'D' => '[tableTwo.d]',
                    ],
                    5 => [
                        'A' => '[tableThree.a]',
                        'C' => null,
                        'D' => '[tableThree.d]',
                    ],
                ],

                'data' => [
                    'tableOne' => [
                        'a' => ['one-a-1'],
                        'd' => ['one-d-1'],
                    ],
                    'tableTwo' => [],
                    'tableThree' => [
                        'a' => ['three-a-1', 'three-a-2', 'three-a-3'],
                        'd' => ['three-d-1', 'three-d-2', 'three-d-3'],
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'one-a-1', 'D' => 'one-d-1'],
                        3 => ['A' => null, 'D' => null],
                        5 => ['A' => 'three-a-1', 'D' => 'three-d-1'],
                        6 => ['A' => 'three-a-2', 'D' => 'three-d-2'],
                        7 => ['A' => 'three-a-3', 'D' => 'three-d-3'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 6, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A5', 'to' => 'A6:A7'],
                        ['from' => 'B5', 'to' => 'B6:B7'],
                        ['from' => 'C5', 'to' => 'C6:C7'],
                        ['from' => 'D5', 'to' => 'D6:D7'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_several_tables3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[tableOne.a]',
                        'B' => '[tableOne.b]',
                    ],
                    3 => [
                        'A' => '[tableTwo.a]',
                        'B' => '[tableTwo.b]',
                    ],
                    5 => [
                        'A' => '[tableThree.a]',
                        'B' => '[tableThree.b]',
                    ],
                ],

                'data' => [
                    'tableOne' => [
                        'a' => ['one-a-1', 'one-a-2', 'one-a-3'],
                        'b' => ['one-b-1', 'one-b-2', 'one-b-3'],
                    ],
                    'tableTwo' => [
                        'a' => ['two-a-1', 'two-a-2', 'two-a-3'],
                        'b' => ['two-b-1', 'two-b-2', 'two-b-3'],
                    ],
                    'tableThree' => [
                        'a' => ['three-a-1', 'three-a-2', 'three-a-3'],
                        'b' => ['three-b-1', 'three-b-2', 'three-b-3'],
                    ],
                ],

                'merge_cells' => ['A1:A2', 'B1:B2', 'A3:A4', 'B3:B4', 'A5:A6', 'B5:B6'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'one-a-1', 'B' => 'one-b-1'],
                        3 => ['A' => 'one-a-2', 'B' => 'one-b-2'],
                        5 => ['A' => 'one-a-3', 'B' => 'one-b-3'],
                        7 => ['A' => 'two-a-1', 'B' => 'two-b-1'],
                        9 => ['A' => 'two-a-2', 'B' => 'two-b-2'],
                        11 => ['A' => 'two-a-3', 'B' => 'two-b-3'],
                        13 => ['A' => 'three-a-1', 'B' => 'three-b-1'],
                        15 => ['A' => 'three-a-2', 'B' => 'three-b-2'],
                        17 => ['A' => 'three-a-3', 'B' => 'three-b-3'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3, 'qty' => 4],
                        ['action' => 'add', 'row' => 9, 'qty' => 4],
                        ['action' => 'add', 'row' => 15, 'qty' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A3'],
                        ['from' => 'A1', 'to' => 'A5'],
                        ['from' => 'A7', 'to' => 'A9'],
                        ['from' => 'A7', 'to' => 'A11'],
                        ['from' => 'A13', 'to' => 'A15'],
                        ['from' => 'A13', 'to' => 'A17'],
                        ['from' => 'B1', 'to' => 'B3'],
                        ['from' => 'B1', 'to' => 'B5'],
                        ['from' => 'B7', 'to' => 'B9'],
                        ['from' => 'B7', 'to' => 'B11'],
                        ['from' => 'B13', 'to' => 'B15'],
                        ['from' => 'B13', 'to' => 'B17'],
                    ],

                    'merge_cells' => [
                        'A3:A4', 'B3:B4',
                        'A5:A6', 'B5:B6',
                        'A9:A10', 'B9:B10',
                        'A11:A12', 'B11:B12',
                        'A15:A16', 'B15:B16',
                        'A17:A18', 'B17:B18',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_several_tables4()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[tableOne.a] [= tableOne]',
                        'B' => '[tableOne.b]',
                    ],
                    3 => [
                        'A' => '[tableTwo.a] [= tableTwo]',
                        'B' => '[tableTwo.b]',
                    ],
                    4 => [
                        'A' => '[tableThree.a] [= tableThree]',
                        'B' => '[tableThree.b]',
                    ],
                ],

                'data' => [
                    'tableOne' => [
                        'a' => ['one-a-1', 'one-a-2', 'one-a-3'],
                        'b' => ['one-b-1', 'one-b-2', 'one-b-3'],
                    ],
                    'tableTwo' => [],
                    'tableThree' => [
                        'a' => ['three-a-1', 'three-a-2', 'three-a-3'],
                        'b' => ['three-b-1', 'three-b-2', 'three-b-3'],
                    ],
                ],

                'merge_cells' => ['A1:A2', 'B1:B2', 'A4:A5', 'B4:B5'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'one-a-1', 'B' => 'one-b-1'],
                        3 => ['A' => 'one-a-2', 'B' => 'one-b-2'],
                        5 => ['A' => 'one-a-3', 'B' => 'one-b-3'],
                        7 => ['A' => 'three-a-1', 'B' => 'three-b-1'],
                        9 => ['A' => 'three-a-2', 'B' => 'three-b-2'],
                        11 => ['A' => 'three-a-3', 'B' => 'three-b-3'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3, 'qty' => 4],
                        ['action' => 'delete', 'row' => 7, 'qty' => 1],
                        ['action' => 'add', 'row' => 9, 'qty' => 4],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A3'],
                        ['from' => 'A1', 'to' => 'A5'],
                        ['from' => 'A7', 'to' => 'A9'],
                        ['from' => 'A7', 'to' => 'A11'],
                        ['from' => 'B1', 'to' => 'B3'],
                        ['from' => 'B1', 'to' => 'B5'],
                        ['from' => 'B7', 'to' => 'B9'],
                        ['from' => 'B7', 'to' => 'B11'],
                    ],

                    'merge_cells' => [
                        'A3:A4', 'B3:B4',
                        'A5:A6', 'B5:B6',
                        'A9:A10', 'B9:B10',
                        'A11:A12', 'B11:B12',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_merge1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[table.a] / [table.b] [$= table.a]',
                        'B' => 'foo',
                        'D' => 'bar',
                    ],
                    2 => [
                        'A' => 'baz',
                        'D' => null,
                    ],
                ],

                'data' => [
                    'table' => [
                        'a' => ['one-a-1', null, 'one-a-3'],
                        'b' => ['one-b-1', 'one-b-2', 'one-b-3'],
                    ],
                ],

                'merge_cells' => ['B1:C1', 'D1:F1'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'one-a-1 / one-b-1', 'B' => 'foo', 'D' => 'bar'],
                        2 => ['A' => null, 'B' => 'foo', 'D' => 'bar'],
                        3 => ['A' => 'one-a-3 / one-b-3', 'B' => 'foo', 'D' => 'bar'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'D1', 'to' => 'D2:D3'],
                    ],

                    'merge_cells' => [
                        'B2:C2', 'D2:F2',
                        'B3:C3', 'D3:F3',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_merge2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'Hello',
                        'B' => '[table.a]',
                        'C' => 'bar 1',
                        'D' => 'bar 2',
                        'F' => 'bar 3',
                    ],
                ],

                'data' => [
                    'table' => [
                        'a' => ['one', 'two', 'three'],
                    ],
                ],

                'merge_cells' => ['A1:A3', 'D1:E1', 'F1:H1'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'Hello', 'B' => 'one', 'C' => 'bar 1', 'D' => 'bar 2', 'F' => 'bar 3'],
                        2 => ['B' => 'two', 'C' => 'bar 1', 'D' => 'bar 2', 'F' => 'bar 3'],
                        3 => ['B' => 'three', 'C' => 'bar 1', 'D' => 'bar 2', 'F' => 'bar 3'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'D1', 'to' => 'D2:D3'],
                        ['from' => 'F1', 'to' => 'F2:F3'],
                    ],

                    'merge_cells' => [
                        'D2:E2', 'F2:H2',
                        'D3:E3', 'F3:H3',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_merge3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'Hello',
                        'C' => '[table.a]',
                        'D' => 0,
                        'E' => '0',
                        'F' => null,
                        'G' => '',
                        'H' => '=ROW()-5',
                    ],
                    2 => [
                        'A' => '[foo]',
                        'B' => '[bar]',
                        'C' => 1,
                        'D' => '1',
                        'H' => null,
                    ],
                ],

                'data' => [
                    'table' => [
                        'a' => ['one', 'two', 'three'],
                    ],
                    'bar' => 'hello',
                ],

                'merge_cells' => ['A1:B3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 'Hello', 'C' => 'one', 'D' => 0, 'E' => '0', 'H' => '=ROW()-5'],
                        2 => ['C' => 'two', 'D' => 0, 'E' => '0', 'H' => '=ROW()-5'],
                        3 => ['C' => 'three', 'D' => 0, 'E' => '0', 'H' => '=ROW()-5'],
                        4 => ['A' => null, 'B' => 'hello'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'D1', 'to' => 'D2:D3'],
                        ['from' => 'E1', 'to' => 'E2:E3'],
                        ['from' => 'F1', 'to' => 'F2:F3'],
                        ['from' => 'G1', 'to' => 'G2:G3'],
                        ['from' => 'H1', 'to' => 'H2:H3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_merge4()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => 'Hello',
                        'B' => null,
                    ],
                    2 => [
                        'B' => '[table.a]',
                    ],
                    4 => [
                        'A' => '[bar]',
                        'B' => null,
                    ],
                ],

                'data' => [
                    'table' => [
                        'a' => ['one', 'two', 'three'],
                    ],
                    'bar' => 'hello',
                ],

                'merge_cells' => ['A1:B3'],
            ],

            [
                'values' => [
                    1 => [
                        'A' => null,
                    ],
                    2 => [
                        'B' => '[table.a]',
                    ],
                    4 => [
                        'A' => '[bar]'
                    ],
                ],

                'data' => [
                    'table' => [
                        'a' => ['one', 'two', 'three'],
                    ],
                    'bar' => 'hello',
                ],

                'merge_cells' => ['A1:B3'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        2 => ['B' => 'one'],
                        3 => ['B' => 'two'],
                        4 => ['B' => 'three'],
                        6 => ['A' => 'hello'],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 3, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'B2', 'to' => 'B3:B4'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_merge5()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(B3:B3)',
                        'B' => null,
                    ],
                    2 => [
                        'A' => 'Foo',
                        'B' => null,
                    ],
                    3 => [
                        'B' => '[product.amount]'
                    ],
                ],

                'data' => [
                    'product' => [
                        'amount' => [111, 222, 333],
                    ],
                ],

                'merge_cells' => ['A2:A4'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => '=SUM(B3:B5)'],
                        3 => ['B' => 111],
                        4 => ['B' => 222],
                        5 => ['B' => 333],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 4, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'B3', 'to' => 'B4:B5'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_merge6()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '=SUM(B3:B3)',
                        'B' => null,
                    ],
                    2 => [
                        'A' => '[product.id]',
                        'B' => null,
                    ],
                    3 => [
                        'B' => '[product.amount]'
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                        'amount' => [111, 222, 333],
                    ],
                ],

                'merge_cells' => ['A2:A4'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => '=SUM(B3:B3)'],
                        2 => ['A' => 1],
                        3 => ['B' => 111],
                        5 => ['A' => 2],
                        6 => ['B' => 222],
                        8 => ['A' => 3],
                        9 => ['B' => 333],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 5, 'qty' => 6],
                    ],

                    'copy_style' => [
                        ['from' => 'A2', 'to' => 'A5'],
                        ['from' => 'A2', 'to' => 'A8'],
                        ['from' => 'B2', 'to' => 'B5'],
                        ['from' => 'B2', 'to' => 'B8'],
                        ['from' => 'B3', 'to' => 'B6'],
                        ['from' => 'B3', 'to' => 'B9'],
                    ],

                    'merge_cells' => ['A5:A7', 'A8:A10'],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_alias1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[product.id]',
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => [],
            ],

            [
                'values' => [
                    1 => [
                        'A' => '[product.id.*]',
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => [],
            ],

            [
                'values' => [
                    1 => [
                        'A' => '[product.id]',
                    ],
                ],

                'data' => [
                    'product' => [
                        ['id' => 1],
                        ['id' => 2],
                        ['id' => 3],
                    ],
                ],

                'merge_cells' => [],
            ],

            [
                'values' => [
                    1 => [
                        'A' => '[product.*.id]',
                    ],
                ],

                'data' => [
                    'product' => [
                        ['id' => 1],
                        ['id' => 2],
                        ['id' => 3],
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 1],
                        2 => ['A' => 2],
                        3 => ['A' => 3],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                    ],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_alias2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[product.*.id]',
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => [],
            ],

            [
                'values' => [
                    1 => [
                        'A' => '[product.id.*]',
                    ],
                ],

                'data' => [
                    'product' => [
                        ['id' => 1],
                        ['id' => 2],
                        ['id' => 3],
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => null],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_alias3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[product.id.2]',
                    ],
                    2 => [
                        'A' => '[product.id.1]',
                    ],
                    3 => [
                        'A' => '[product.id.0]',
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => [],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 3],
                        2 => ['A' => 2],
                        3 => ['A' => 1],
                    ],

                    'rows' => [],

                    'copy_style' => [],

                    'merge_cells' => [],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_styles1()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[product.id]',
                        'E' => null,
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => ['C1:D1'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 1],
                        2 => ['A' => 2],
                        3 => ['A' => 3],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'E1', 'to' => 'E2:E3'],
                    ],

                    'merge_cells' => [
                        'C2:D2', 'C3:D3',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_styles2()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[product.id]',
                        'D' => null,
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => ['C1:D1'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 1],
                        2 => ['A' => 2],
                        3 => ['A' => 3],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                    ],

                    'merge_cells' => [
                        'C2:D2', 'C3:D3',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_styles3()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'B' => '[product.id]',
                        'E' => null,
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => ['C1:D1'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['B' => 1],
                        2 => ['B' => 2],
                        3 => ['B' => 3],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'E1', 'to' => 'E2:E3'],
                    ],

                    'merge_cells' => [
                        'C2:D2', 'C3:D3',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_styles4()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[product.id]',
                        'BA' => null,
                    ],
                ],

                'data' => [
                    'product' => [
                        'id' => [1, 2, 3],
                    ],
                ],

                'merge_cells' => ['AB1:BA1'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['A' => 1],
                        2 => ['A' => 2],
                        3 => ['A' => 3],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'AA1', 'to' => 'AA2:AA3'],
                        ['from' => 'AB1', 'to' => 'AB2:AB3'],
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'D1', 'to' => 'D2:D3'],
                        ['from' => 'E1', 'to' => 'E2:E3'],
                        ['from' => 'F1', 'to' => 'F2:F3'],
                        ['from' => 'G1', 'to' => 'G2:G3'],
                        ['from' => 'H1', 'to' => 'H2:H3'],
                        ['from' => 'I1', 'to' => 'I2:I3'],
                        ['from' => 'J1', 'to' => 'J2:J3'],
                        ['from' => 'K1', 'to' => 'K2:K3'],
                        ['from' => 'L1', 'to' => 'L2:L3'],
                        ['from' => 'M1', 'to' => 'M2:M3'],
                        ['from' => 'N1', 'to' => 'N2:N3'],
                        ['from' => 'O1', 'to' => 'O2:O3'],
                        ['from' => 'P1', 'to' => 'P2:P3'],
                        ['from' => 'Q1', 'to' => 'Q2:Q3'],
                        ['from' => 'R1', 'to' => 'R2:R3'],
                        ['from' => 'S1', 'to' => 'S2:S3'],
                        ['from' => 'T1', 'to' => 'T2:T3'],
                        ['from' => 'U1', 'to' => 'U2:U3'],
                        ['from' => 'V1', 'to' => 'V2:V3'],
                        ['from' => 'W1', 'to' => 'W2:W3'],
                        ['from' => 'X1', 'to' => 'X2:X3'],
                        ['from' => 'Y1', 'to' => 'Y2:Y3'],
                        ['from' => 'Z1', 'to' => 'Z2:Z3'],
                    ],

                    'merge_cells' => [
                        'AB2:BA2', 'AB3:BA3',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_styles5()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'C' => '[foo.id]',
                        'E' => null,
                    ],
                    2 => [
                        'A' => '[bar.id]',
                        'E' => null,
                    ],
                ],

                'data' => [
                    'foo' => [
                        'id' => [1, 2, 3],
                    ],
                    'bar' => [
                        'id' => [4, 5, 6],
                    ],
                ],

                'merge_cells' => ['C1:E1', 'A2:D2'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['C' => 1],
                        2 => ['C' => 2],
                        3 => ['C' => 3],
                        4 => ['A' => 4],
                        5 => ['A' => 5],
                        6 => ['A' => 6],
                    ],

                    'rows' => [
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                        ['action' => 'add', 'row' => 5, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'A4', 'to' => 'A5:A6'],
                        ['from' => 'B1', 'to' => 'B2:B3'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'E4', 'to' => 'E5:E6'],
                    ],

                    'merge_cells' => [
                        'C2:E2', 'C3:E3',
                        'A5:D5', 'A6:D6',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }

    /**
     * @return void
     */
    public function test_schema_styles6()
    {
        $data = [
            [
                'values' => [
                    1 => [
                        'A' => '[= baz]',
                        'H' => null,
                    ],
                    2 => [
                        'C' => '[foo.id]',
                        'H' => null,
                    ],
                    4 => [
                        'A' => '[bar.id]',
                        'H' => null,
                    ],
                ],

                'data' => [
                    'foo' => [
                        'id' => [1, 2, 3],
                    ],
                    'bar' => [
                        'id' => [4, 5, 6],
                    ],
                ],

                'merge_cells' => ['A2:B2', 'C2:G2', 'A4:H4'],
            ],
        ];

        foreach ($data as $id => $item) {
            $this->assertSame(
                [
                    'data' => [
                        1 => ['C' => 1],
                        2 => ['C' => 2],
                        3 => ['C' => 3],
                        5 => ['A' => 4],
                        6 => ['A' => 5],
                        7 => ['A' => 6],
                    ],

                    'rows' => [
                        ['action' => 'delete', 'row' => 1, 'qty' => 1],
                        ['action' => 'add', 'row' => 2, 'qty' => 2],
                        ['action' => 'add', 'row' => 6, 'qty' => 2],
                    ],

                    'copy_style' => [
                        ['from' => 'A1', 'to' => 'A2:A3'],
                        ['from' => 'A5', 'to' => 'A6:A7'],
                        ['from' => 'C1', 'to' => 'C2:C3'],
                        ['from' => 'H1', 'to' => 'H2:H3'],
                    ],

                    'merge_cells' => [
                        'A2:B2', 'C2:G2', 'A3:B3', 'C3:G3',
                        'A6:H6', 'A7:H7',
                    ],

                    'copy_width' => [],
                ],
                $this->service->schema($item['values'], $item['data'], $item['merge_cells'])->toArray(),
                "$id"
            );
        }
    }
}
