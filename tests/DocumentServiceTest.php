<?php

namespace AnourValar\Office\Tests;

use AnourValar\Office\DocumentService;
use AnourValar\Office\Format;
use AnourValar\Office\Generated;

class DocumentServiceTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @return void
     */
    public function test_generate_returns_generated()
    {
        $driver = $this->getDriver();
        $generated = (new DocumentService($driver))->generate('template.docx', ['foo' => 'bar']);

        $this->assertInstanceOf(Generated::class, $generated);
        $this->assertSame($driver, $generated->driver);
    }

    /**
     * Data is flattened into bracketed markers before being passed to the driver.
     *
     * @return void
     */
    public function test_generate_canonizes_data()
    {
        $driver = $this->getDriver();

        (new DocumentService($driver))->generate('template.docx', [
            'title' => 'Hello',
            'user' => [
                'name' => 'John',
                'roles' => ['admin', 'editor'],
            ],
        ]);

        $this->assertSame([
            '[title]' => 'Hello',
            '[user.name]' => 'John',
            '[user.roles.0]' => 'admin',
            '[user.roles.1]' => 'editor',
        ], $driver->replacedData);
    }

    /**
     * @return void
     */
    public function test_generate_supports_to_array_objects()
    {
        $driver = $this->getDriver();

        $data = new class ()
        {
            public function toArray(): array
            {
                return ['foo' => 'bar', 'nested' => ['baz' => 1]];
            }
        };

        (new DocumentService($driver))->generate('template.docx', $data);

        $this->assertSame(['[foo]' => 'bar', '[nested.baz]' => 1], $driver->replacedData);
    }

    /**
     * @return void
     */
    public function test_generate_detects_format_from_extension()
    {
        $driver = $this->getDriver();
        (new DocumentService($driver))->generate('letter.DOCX', []);

        $this->assertSame(Format::Docx, $driver->loadedFormat);
        $this->assertSame('letter.DOCX', $driver->loadedFile);
    }

    /**
     * Unknown extensions fall back to Docx.
     *
     * @return void
     */
    public function test_generate_defaults_to_docx()
    {
        $driver = $this->getDriver();
        (new DocumentService($driver))->generate('template', []);

        $this->assertSame(Format::Docx, $driver->loadedFormat);
    }

    /**
     * @return void
     */
    public function test_generate_accepts_stringable_template()
    {
        $driver = $this->getDriver();

        $template = new class () implements \Stringable
        {
            public function __toString(): string
            {
                return 'from-buffer.docx';
            }
        };

        (new DocumentService($driver))->generate($template, []);

        $this->assertSame(Format::Docx, $driver->loadedFormat);
    }

    /**
     * @return object
     */
    protected function getDriver(): object
    {
        return new class () implements \AnourValar\Office\Drivers\DocumentInterface
        {
            public ?array $replacedData = null;
            public ?Format $loadedFormat = null;
            public ?string $loadedFile = null;

            public function load(string $file, Format $format): \AnourValar\Office\Drivers\SaveInterface
            {
                $this->loadedFile = $file;
                $this->loadedFormat = $format;

                return $this;
            }

            public function save(string $file, Format $format): void
            {
                //
            }

            public function replace(array $data): self
            {
                $this->replacedData = $data;

                return $this;
            }
        };
    }
}
