<?php

namespace AnourValar\Office\Drivers;

interface SaveInterface
{
    /**
     * Save in specific format
     *
     * @param string $file
     * @param \AnourValar\Office\Format $format
     * @return void
     */
    public function save(string $file, \AnourValar\Office\Format $format): void;
}
