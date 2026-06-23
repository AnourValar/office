<?php

namespace AnourValar\Office\Tests;

class TraitsTest extends \PHPUnit\Framework\TestCase
{
    use \AnourValar\Office\Traits\Parser;

    /**
     * @return void
     */
    public function test_isColumnLE()
    {
        $this->assertTrue($this->isColumnLE('A', 'B'));
        $this->assertTrue($this->isColumnLE('B', 'B'));
        $this->assertFalse($this->isColumnLE('C', 'B'));

        $this->assertTrue($this->isColumnLE('Z', 'AA'));
        $this->assertTrue($this->isColumnLE('AA', 'AA'));
        $this->assertFalse($this->isColumnLE('AB', 'AA'));

        $this->assertFalse($this->isColumnLE('AAA', 'ZZ'));
        $this->assertTrue($this->isColumnLE('AAA', 'AAA'));
        $this->assertTrue($this->isColumnLE('AAA', 'AAB'));
    }

    /**
     * @return void
     */
    public function test_isColumnGE()
    {
        $this->assertFalse($this->isColumnGE('A', 'B'));
        $this->assertTrue($this->isColumnGE('B', 'B'));
        $this->assertTrue($this->isColumnGE('C', 'B'));

        $this->assertFalse($this->isColumnGE('Z', 'AA'));
        $this->assertTrue($this->isColumnGE('AA', 'AA'));
        $this->assertTrue($this->isColumnGE('AB', 'AA'));

        $this->assertTrue($this->isColumnGE('AAA', 'ZZ'));
        $this->assertTrue($this->isColumnGE('AAA', 'AAA'));
        $this->assertFalse($this->isColumnGE('AAA', 'AAB'));
    }

    /**
     * @return void
     */
    public function test_sort()
    {
        $columns = ['BA' => 7,  'AZ' => 6, 'A' => 1, 'B' => 2, 'Z' => 3, 'AA' => 4, 'AB' => 5];

        uksort($columns, fn ($a, $b) => $this->isColumnLE($a, $b) ? -1 : 1);

        $this->assertSame(
            ['A' => 1, 'B' => 2, 'Z' => 3, 'AA' => 4, 'AB' => 5, 'AZ' => 6, 'BA' => 7],
            $columns
        );
    }

    /**
     * @return void
     */
    public function test_dot()
    {
        $this->assertSame(
            ['a' => 1, 'b.c' => 2, 'b.d.e' => 3],
            $this->dot(['a' => 1, 'b' => ['c' => 2, 'd' => ['e' => 3]]])
        );

        // empty arrays produce no keys
        $this->assertSame(['x' => 1], $this->dot(['x' => 1, 'y' => []]));

        // numeric keys are preserved as-is
        $this->assertSame(
            ['list.0' => 'foo', 'list.1' => 'bar'],
            $this->dot(['list' => ['foo', 'bar']])
        );

        $this->assertSame([], $this->dot([]));
    }

    /**
     * @return void
     */
    public function test_dot_with_prefix()
    {
        $this->assertSame(
            ['root.a' => 1, 'root.b.c' => 2],
            $this->dot(['a' => 1, 'b' => ['c' => 2]], 'root.')
        );
    }

    /**
     * @return void
     */
    public function test_strIncrement()
    {
        $this->assertSame('B', $this->strIncrement('A'));
        $this->assertSame('AA', $this->strIncrement('Z'));
        $this->assertSame('AB', $this->strIncrement('AA'));
        $this->assertSame('BA', $this->strIncrement('AZ'));
        $this->assertSame('AAA', $this->strIncrement('ZZ'));
    }
}
