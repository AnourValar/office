<?php

namespace AnourValar\Office\Tests;

use AnourValar\Office\Format;
use AnourValar\Office\Generated;
use AnourValar\Office\Mixer;

class MixerTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @return void
     */
    public function test_merges_sheets_into_reference_driver()
    {
        $reference = $this->getDriver(['Main']);
        $second = $this->getDriver(['Main', 'Extra']);

        $result = (new Mixer())(new Generated($reference), new Generated($second));

        $this->assertInstanceOf(Generated::class, $result);
        $this->assertSame($reference, $result->driver);

        // colliding titles are deduplicated, unique ones are kept
        $this->assertSame(['Main (1)', 'Extra'], $second->titles);

        // each merged sheet is unioned into the reference driver
        $this->assertSame(['Main (1)', 'Extra'], $reference->merged);
    }

    /**
     * @return void
     */
    public function test_merges_multiple_documents_with_progressive_dedup()
    {
        $reference = $this->getDriver(['Sheet']);
        $a = $this->getDriver(['Sheet']);
        $b = $this->getDriver(['Sheet']);

        (new Mixer())(new Generated($reference), new Generated($a), new Generated($b));

        $this->assertSame(['Sheet (1)'], $a->titles);
        $this->assertSame(['Sheet (2)'], $b->titles);
        $this->assertSame(['Sheet (1)', 'Sheet (2)'], $reference->merged);
    }

    /**
     * @return void
     */
    public function test_single_document_is_returned_untouched()
    {
        $reference = $this->getDriver(['Main', 'Second']);

        $result = (new Mixer())(new Generated($reference));

        $this->assertSame($reference, $result->driver);
        $this->assertSame([], $reference->merged);
        $this->assertSame(['Main', 'Second'], $reference->titles);
    }

    /**
     * @return void
     */
    public function test_throws_when_first_argument_is_not_generated()
    {
        $this->expectException(\LogicException::class);
        $this->expectExceptionMessage('Input data must be instanceof Generated');

        (new Mixer())('not-generated');
    }

    /**
     * @return void
     */
    public function test_throws_when_driver_does_not_implement_mix_interface()
    {
        $driver = new class () implements \AnourValar\Office\Drivers\SaveInterface
        {
            public function save(string $file, Format $format): void
            {
                //
            }
        };

        $this->expectException(\LogicException::class);
        $this->expectExceptionMessage('Driver must implements MixInterface.');

        (new Mixer())(new Generated($driver));
    }

    /**
     * @return void
     */
    public function test_throws_when_subsequent_argument_is_not_generated()
    {
        $reference = $this->getDriver(['Main']);

        $this->expectException(\LogicException::class);
        $this->expectExceptionMessage('Input data must be instanceof Generated');

        (new Mixer())(new Generated($reference), 'broken');
    }

    /**
     * @return void
     */
    public function test_throws_when_drivers_are_different_implementations()
    {
        $reference = $this->getDriver(['Main']);
        $other = $this->getOtherDriver();

        $this->expectException(\LogicException::class);
        $this->expectExceptionMessage('All drivers should be instances of the same implementation.');

        (new Mixer())(new Generated($reference), new Generated($other));
    }

    /**
     * Each call returns an instance of the same anonymous class, so the Mixer's
     * same-implementation check treats them as compatible drivers.
     *
     * @param array $titles
     * @return object
     */
    protected function getDriver(array $titles): object
    {
        return new class ($titles) implements \AnourValar\Office\Drivers\MixInterface, \AnourValar\Office\Drivers\SaveInterface
        {
            public array $merged = [];
            protected int $active = 0;

            public function __construct(public array $titles)
            {
                //
            }

            public function setSheet(int $index): self
            {
                $this->active = $index;

                return $this;
            }

            public function getSheetCount(): int
            {
                return count($this->titles);
            }

            public function getSheetTitle(): string
            {
                return $this->titles[$this->active];
            }

            public function setSheetTitle(string $title): self
            {
                $this->titles[$this->active] = $title;

                return $this;
            }

            public function mergeDriver(\AnourValar\Office\Drivers\MixInterface $driver): self
            {
                $this->merged[] = $driver->getSheetTitle();

                return $this;
            }

            public function save(string $file, Format $format): void
            {
                //
            }
        };
    }

    /**
     * A distinct (unrelated) implementation to trigger the "same implementation" guard.
     *
     * @return object
     */
    protected function getOtherDriver(): object
    {
        return new class () implements \AnourValar\Office\Drivers\MixInterface, \AnourValar\Office\Drivers\SaveInterface
        {
            public function setSheet(int $index): self
            {
                return $this;
            }

            public function getSheetCount(): int
            {
                return 1;
            }

            public function getSheetTitle(): string
            {
                return 'Main';
            }

            public function setSheetTitle(string $title): self
            {
                return $this;
            }

            public function mergeDriver(\AnourValar\Office\Drivers\MixInterface $driver): self
            {
                return $this;
            }

            public function save(string $file, Format $format): void
            {
                //
            }
        };
    }
}
