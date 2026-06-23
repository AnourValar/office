<?php

namespace AnourValar\Office\Tests;

use AnourValar\Office\Buffer;

class BufferTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @return void
     */
    public function test_creates_temporary_file_with_content()
    {
        $buffer = new Buffer('hello world');

        $filename = (string) $buffer;

        $this->assertFileExists($filename);
        $this->assertSame('hello world', file_get_contents($filename));
    }

    /**
     * @return void
     */
    public function test_supports_binary_content()
    {
        $binary = random_bytes(256);
        $buffer = new Buffer($binary);

        $this->assertSame($binary, file_get_contents((string) $buffer));
    }

    /**
     * @return void
     */
    public function test_empty_buffer()
    {
        $buffer = new Buffer('');

        $this->assertFileExists((string) $buffer);
        $this->assertSame('', file_get_contents((string) $buffer));
    }

    /**
     * The underlying temporary file is removed once the buffer is destructed.
     *
     * @return void
     */
    public function test_file_removed_on_destruct()
    {
        $buffer = new Buffer('temp');
        $filename = (string) $buffer;
        $this->assertFileExists($filename);

        unset($buffer);

        $this->assertFileDoesNotExist($filename);
    }
}
