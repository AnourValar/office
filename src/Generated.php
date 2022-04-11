<?php

namespace AnourValar\Office;

class Generated
{
    /**
     * @var \AnourValar\Office\Drivers\SaveInterface
     */
    public readonly \AnourValar\Office\Drivers\SaveInterface $driver;

    /**
     * Handle template's saving
     *
     * @var \Closure(\AnourValar\Office\Drivers\SaveInterface $driver, \AnourValar\Office\Format $format)|null
     */
    protected ?\Closure $hookSave = null;

    /**
     * @param \AnourValar\Office\Drivers\SaveInterface $driver
     * @return void
     */
    public function __construct(\AnourValar\Office\Drivers\SaveInterface $driver)
    {
        $this->driver = $driver;
    }

    /**
     * @see magic
     *
     * @return string
     */
    public function __toString()
    {
        return $this->save();
    }

    /**
     * Save generated document to the buffer
     *
     * @param \AnourValar\Office\Format $format
     * @return string
     */
    public function save(Format $format = Format::Xlsx): string
    {
        ob_start();

        if ($this->hookSave) {
            ($this->hookSave)($this->driver, $format);
        } else {
            $this->driver->save('php://output', $format);
        }

        return ob_get_clean();
    }

    /**
     * Save generated document to the file
     *
     * @param string $filename
     * @param \AnourValar\Office\Format $format
     * @return int|NULL
     */
    public function saveAs(string $filename, Format $format = Format::Xlsx): ?int
    {
        return file_put_contents($filename, $this->save($format));
    }

    /**
     * Set hookSave
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookSave(?\Closure $closure): self
    {
        $this->hookSave = $closure;

        return $this;
    }
}
