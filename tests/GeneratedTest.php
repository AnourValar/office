<?php

namespace AnourValar\Office\Tests;

use AnourValar\Office\Format;
use AnourValar\Office\Generated;

class GeneratedTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @return void
     */
    public function test_save_returns_buffer_from_driver()
    {
        $generated = new Generated($this->getDriver());

        $this->assertSame('saved:php://output:xlsx', $generated->save(Format::Xlsx));
        $this->assertSame('saved:php://output:pdf', $generated->save(Format::Pdf));
    }

    /**
     * @return void
     */
    public function test_saveAs_writes_file_and_returns_bytes()
    {
        $generated = new Generated($this->getDriver());
        $filename = tempnam(sys_get_temp_dir(), 'office_test_') . '.xlsx';

        try {
            $bytes = $generated->saveAs($filename);

            $this->assertIsInt($bytes);
            $this->assertSame('saved:php://output:xlsx', file_get_contents($filename));
            $this->assertSame(strlen('saved:php://output:xlsx'), $bytes);
        } finally {
            @unlink($filename);
        }
    }

    /**
     * Format is inferred from the file extension when not given explicitly.
     *
     * @return void
     */
    public function test_saveAs_infers_format_from_extension()
    {
        $generated = new Generated($this->getDriver());
        $filename = tempnam(sys_get_temp_dir(), 'office_test_') . '.pdf';

        try {
            $generated->saveAs($filename);

            $this->assertSame('saved:php://output:pdf', file_get_contents($filename));
        } finally {
            @unlink($filename);
        }
    }

    /**
     * @return void
     */
    public function test_saveAs_with_explicit_format_overrides_extension()
    {
        $generated = new Generated($this->getDriver());
        $filename = tempnam(sys_get_temp_dir(), 'office_test_') . '.bin';

        try {
            $generated->saveAs($filename, Format::Csv);

            $this->assertSame('saved:php://output:csv', file_get_contents($filename));
        } finally {
            @unlink($filename);
        }
    }

    /**
     * @return void
     */
    public function test_hookSave_overrides_default_save()
    {
        $captured = null;
        $generated = (new Generated($this->getDriver()))
            ->hookSave(function ($driver, Format $format) use (&$captured) {
                $captured = $format;
                echo 'hooked!';
            });

        $this->assertSame('hooked!', $generated->save(Format::Ods));
        $this->assertSame(Format::Ods, $captured);
    }

    /**
     * @return void
     */
    public function test_hookSave_is_fluent_and_resettable()
    {
        $generated = new Generated($this->getDriver());

        $this->assertSame($generated, $generated->hookSave(fn () => null));

        // resetting back to null restores the default driver-based save
        $generated->hookSave(null);
        $this->assertSame('saved:php://output:html', $generated->save(Format::Html));
    }

    /**
     * @return \AnourValar\Office\Drivers\SaveInterface
     */
    protected function getDriver(): \AnourValar\Office\Drivers\SaveInterface
    {
        return new class () implements \AnourValar\Office\Drivers\SaveInterface
        {
            public function save(string $file, Format $format): void
            {
                echo "saved:{$file}:{$format->value}";
            }
        };
    }
}
