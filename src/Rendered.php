<?php

namespace AnourValar\Office;

class Rendered
{
    /**
     * Binary data
     *
     * @var string
     */
    private string $payload;

    /**
     * @param string $payload
     * @return void
     */
    public function __construct(string $payload)
    {
        $this->payload = $payload;
    }

    /**
     * @see magic
     *
     * @return string
     */
    public function __toString()
    {
        return $this->payload;
    }

    /**
     * Save rendered data to file
     *
     * @param string $filename
     * @return int|NULL
     */
    public function save(string $filename): ?int
    {
        return file_put_contents($filename, $this->payload);
    }
}
