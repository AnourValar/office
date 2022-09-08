<?php

namespace AnourValar\Office\Drivers;

interface DocumentInterface extends SaveInterface, LoadInterface
{
    /**
     * Replace markers with values
     *
     * @param array $data
     * @return self
     */
    public function replace(array $data): self;
}
