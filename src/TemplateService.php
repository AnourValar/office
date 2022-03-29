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
     * @param string $template
     * @param mixed $data
     * @param \AnourValar\Office\SaveFormat $saveFormat
     * @throws \LogicException
     * @return \AnourValar\Office\Rendered
     */
    public function render(string $template, mixed $data, SaveFormat $saveFormat = SaveFormat::Xlsx): Rendered
    {
        // Get instance of driver
        $driver = new $this->driverClass();
        if (! $driver instanceof \AnourValar\Office\Drivers\TemplateInterface) {
            throw new \LogicException('Driver must implements TemplateInterface.');
        }

        // Handle with input data
        $data = $this->parser->canonizeData($data);

        // Open the template
        $driver->loadXlsx($template);

        // Hook: before
        if ($this->hookBefore) {
            ($this->hookBefore)($driver, $data);
        }

        // Get schema of the document
        $schema = $this->parser->schema($driver->getValues(null), $data);

        // Rows add/delete
        foreach ($schema['rows'] as $row) {
            if ($row['action'] == 'add') {
                $driver->addRow($row['row']);
            } elseif ($row['action'] == 'delete') {
                $driver->deleteRow($row['row']);
            } else {
                throw new \LogicException('Incorrect usage.');
            }
        }

        // Copy styles
        foreach ($schema['copy_styles'] as $item) {
            $driver->copyStyle($item['from'], $item['to']);
        }

        // Set width
        foreach ($schema['width'] as $columnTo => $columnFrom) {
            $driver->setWidth($columnTo, $columnFrom);
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
        $driver->{$saveFormat->value}('php://output');
        return new Rendered(ob_get_clean());
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
