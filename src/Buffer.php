<?php

namespace AnourValar\Office;

class Buffer implements \Stringable
{
    /**
     * @var resource
     */
    protected readonly mixed $resource;

    /**
     * @var string
     */
    protected readonly string $filename;

    /**
     * Creates a temporary file from the buffer
     *
     * @param string $buffer
     * @return void
     */
    public function __construct(string $buffer)
    {
        $this->resource = tmpfile();
        fwrite($this->resource, $buffer);

        $this->filename = stream_get_meta_data($this->resource)['uri'];
    }

    /**
     * @see magic
     *
     * @return string
     */
    public function __toString(): string
    {
        return $this->filename;
    }

    /**
     * @return void
     */
    public function __destruct()
    {
        fclose($this->resource);
    }
}
