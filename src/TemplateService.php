<?php

namespace AnourValar\Office;

class TemplateService
{
    /**
     * @var string
     */
    protected string $driverClass;

    /**
     * @var \AnourValar\Office\Template\Parser
     */
    protected \AnourValar\Office\Template\Parser $parser;

    /**
     * Handle template's loading
     *
     * @var \Closure(TemplateInterface $templateDriver, string $templateFile)|null
     */
    protected ?\Closure $hookLoad = null;

    /**
     * Actions with template before data inserted
     *
     * @var \Closure(TemplateInterface $templateDriver, array &$data)|null
     */
    protected ?\Closure $hookBefore = null;

    /**
     * Cell's value handler (on set)
     *
     * @var \Closure(TemplateInterface $templateDriver, string $cell, mixed $value)|null
     */
    protected ?\Closure $hookValue = null;

    /**
     * Actions with template after data inserted
     *
     * @var \Closure(TemplateInterface $templateDriver)|null
     */
    protected ?\Closure $hookAfter = null;

    /**
     * Handle template's saving
     *
     * @var \Closure(TemplateInterface $templateDriver)|null
     */
    protected ?\Closure $hookSave = null;

    /**
     * @param string $driverClass
     * @param \AnourValar\Office\Template\Parser $parser
     * @return void
     */
    public function __construct(
        string $driverClass = \AnourValar\Office\Drivers\PhpSpreadsheetDriver::class,
        \AnourValar\Office\Template\Parser $parser = new \AnourValar\Office\Template\Parser()
    ) {
        $this->driverClass = $driverClass;
        $this->parser = $parser;
    }

    /**
     * @param string $templateFile
     * @param mixed $data
     * @param \AnourValar\Office\SaveFormat $saveFormat
     * @throws \LogicException
     * @return \AnourValar\Office\Rendered
     */
    public function render(string $templateFile, mixed $data, SaveFormat $saveFormat = SaveFormat::Xlsx): Rendered
    {
        // Get instance of driver
        $driver = new $this->driverClass();
        if (! $driver instanceof \AnourValar\Office\Drivers\TemplateInterface) {
            throw new \LogicException('Driver must implements TemplateInterface.');
        }

        // Handle with input data
        $data = $this->parser->canonizeData($data);

        // Open the template
        if ($this->hookLoad) {
            ($this->hookLoad)($driver, $templateFile);
        } else {
            $driver->loadXlsx($templateFile);
        }

        // Hook: before
        if ($this->hookBefore) {
            ($this->hookBefore)($driver, $data);
        }

        // Get schema of the document
        $schema = $this->parser->schema($driver->getValues(null), $data, $driver->getMergeCells())->toArray();

        // rows
        foreach ($schema['rows'] as $row) {
            if ($row['action'] == 'add') {
                $driver->addRow($row['row']);
            } elseif ($row['action'] == 'delete') {
                $driver->deleteRow($row['row']);
            } else {
                throw new \LogicException('Incorrect usage.');
            }
        }

        // copy_style
        foreach ($schema['copy_style'] as $item) {
            $driver->copyStyle($item['from'], $item['to']);
        }

        // merge_cells
        foreach ($schema['merge_cells'] as $item) {
            $driver->mergeCells($item);
        }

        // copy_width
        foreach ($schema['copy_width'] as $item) {
            $driver->setWidth($item['to'], $item['from']);
        }

        // Decode data
        $schema = $this->handleValue($schema, $driver);

        // Replace markers (and last but not least :D)
        $driver->setValues($schema['data']);

        // Hook: after
        if ($this->hookAfter) {
            ($this->hookAfter)($driver);
        }

        // Render & save to the buffer
        ob_start();
        if ($this->hookSave) {
            ($this->hookSave)($driver);
        } else {
            $driver->{$saveFormat->driverSaveMethod()}('php://output');
        }
        return new Rendered(ob_get_clean());
    }

    /**
     * Set hookLoad
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookLoad(?\Closure $closure): self
    {
        $this->hookLoad = $closure;

        return $this;
    }

    /**
     * Set hookBefore
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookBefore(?\Closure $closure): self
    {
        $this->hookBefore = $closure;

        return $this;
    }

    /**
     * Set hookValue
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookValue(?\Closure $closure): self
    {
        $this->hookValue = $closure;

        return $this;
    }

    /**
     * Set hookAfter
     *
     * @param ?\Closure $closure
     * @return self
     */
    public function hookAfter(?\Closure $closure): self
    {
        $this->hookAfter = $closure;

        return $this;
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

    /**
     * Set driver class
     *
     * @param string $driverClass
     * @return self
     */
    public function setDriverClass(string $driverClass): self
    {
        $this->driverClass = $driverClass;

        return $this;
    }

    /**
     * @param array $schema
     * @param \AnourValar\Office\Drivers\TemplateInterface $driver
     * @return array
     */
    protected function handleValue(array $schema, \AnourValar\Office\Drivers\TemplateInterface $driver): array
    {
        foreach ($schema['data'] as $row => &$columns) {
            foreach ($columns as $column => &$value) {
                if (is_null($value)) {
                    continue;
                }

                if ($value instanceof \Closure) {
                    // Private Closure
                    $value = $value($driver, $column.$row);
                } elseif ($this->hookValue) {
                    // Hook: value
                    $value = ($this->hookValue)($driver, $column.$row, $value);
                }

                if (is_null($value)) {
                    unset($schema['data'][$row][$column]);
                }
            }
            unset($value);
        }
        unset($columns);

        return $schema;
    }
}
