<?php

namespace AnourValar\Office\Drivers;

interface LoadInterface
{
    /**
     * Load a template with specific format
     *
     * @param string $file
     * @param \AnourValar\Office\Format $format
     * @return self
     */
    public function load(string $file, \AnourValar\Office\Format $format): self;
}
