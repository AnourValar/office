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
}
