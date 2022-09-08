<?php

namespace AnourValar\Office\Drivers;

interface LoadInterface
{
    /**
     * Load a template with specific format
     *
     * @param string $file
     * @param \AnourValar\Office\Format $format
     * @return \AnourValar\Office\Drivers\SaveInterface
     */
    public function load(string $file, \AnourValar\Office\Format $format): SaveInterface;
}
