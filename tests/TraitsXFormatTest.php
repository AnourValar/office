<?php

namespace AnourValar\Office\Tests;

class TraitsXFormatTest extends \PHPUnit\Framework\TestCase
{
    use \AnourValar\Office\Traits\XFormat;

    /**
     * Excel stores dates as a serial number of days since 1899-12-30.
     *
     * @return void
     */
    public function test_excelDate()
    {
        $utc = new \DateTimeZone('UTC');

        $this->assertSame(1.0, $this->excelDate(new \DateTime('1900-01-01', $utc)));
        $this->assertSame(59.0, $this->excelDate(new \DateTime('1900-02-28', $utc)));
        // Excel's famous Lotus 1-2-3 leap year bug: serial 60 (1900-02-29) is skipped here
        $this->assertSame(61.0, $this->excelDate(new \DateTime('1900-03-01', $utc)));
        $this->assertSame(44197.0, $this->excelDate(new \DateTime('2021-01-01', $utc)));
    }

    /**
     * @return void
     */
    public function test_excelDate_with_time()
    {
        $utc = new \DateTimeZone('UTC');

        $this->assertSame(45351.5, $this->excelDate(new \DateTime('2024-02-29 12:00:00', $utc)));
        $this->assertSame(45824.25, $this->excelDate(new \DateTime('2025-06-16 06:00:00', $utc)));

        // midnight has no fractional part
        $this->assertSame(45824.0, $this->excelDate(new \DateTime('2025-06-16 00:00:00', $utc)));
    }

    /**
     * @return void
     */
    public function test_escape()
    {
        $this->assertSame('a&lt;b&gt;', $this->escape('a<b>'));
        $this->assertSame('&amp;', $this->escape('&'));
        $this->assertSame('&quot;', $this->escape('"'));
        $this->assertSame('&#039;', $this->escape("'"));
        $this->assertSame('plain text', $this->escape('plain text'));
    }

    /**
     * @return void
     */
    public function test_escape_null_becomes_empty_string()
    {
        $this->assertSame('', $this->escape(null));
    }

    /**
     * Entities are encoded unconditionally (double_encode = true).
     *
     * @return void
     */
    public function test_escape_double_encodes()
    {
        $this->assertSame('&amp;amp;', $this->escape('&amp;'));
        $this->assertSame('&amp;lt;b&amp;gt;', $this->escape('&lt;b&gt;'));
    }
}
