<?php

namespace AnourValar\Office;

class MixerService
{
    /**
     * Mix generated documents
     *
     * @param \AnourValar\Office\Generated[\AnourValar\Office\Drivers\MixInterface] $generated
     * @throws \LogicException
     * @return \AnourValar\Office\Generated
     */
    public function __invoke(...$generated): Generated
    {
        $referenceDriver = array_shift($generated);
        if (! $referenceDriver instanceof \AnourValar\Office\Generated) {
            throw new \LogicException('Input data must be instanceof Generated');
        }

        $referenceDriver = $referenceDriver->driver;
        if (! $referenceDriver instanceof \AnourValar\Office\Drivers\MixInterface) {
            throw new \LogicException('Driver must implements MixInterface.');
        }

        $titles = [ $referenceDriver->getSheetTitle() ];

        foreach ($generated as $driver) {
            if (! $driver instanceof \AnourValar\Office\Generated) {
                throw new \LogicException('Input data must be instanceof Generated');
            }

            $driver = $driver->driver;
            if (! $driver instanceof $referenceDriver) {
                throw new \LogicException('All drivers must instances of the same implementation.');
            }

            $driver->setSheetTitle( $titles[] = $this->getTitle($driver->getSheetTitle(), $titles) );

            $referenceDriver->mergeDriver($driver);
        }

        return new Generated($referenceDriver);
    }

    /**
     * @param string $title
     * @param array $titles
     * @return string
     */
    protected function getTitle(string $title, array $titles): string
    {
        while (in_array($title, $titles, true)) {
            $title = preg_replace_callback(
                '#\((\d+)\)$#',
                fn ($patterns) => '(' . ++$patterns[1] . ')',
                $title,
                -1,
                $count
            );

            if (! $count) {
                $title .= ' (1)';
            }

        }

        return  $title;
    }
}
