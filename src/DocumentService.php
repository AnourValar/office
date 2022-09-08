<?php

namespace AnourValar\Office;

use AnourValar\Office\Drivers\DocumentInterface;

class DocumentService
{
    use \AnourValar\Office\Traits\Parser;

    /**
     * @var \AnourValar\Office\Drivers\DocumentInterface
     */
    protected \AnourValar\Office\Drivers\DocumentInterface $driver;

    /**
     * @param \AnourValar\Office\Drivers\DocumentInterface $driver
     * @return void
     */
    public function __construct(DocumentInterface $driver = new \AnourValar\Office\Drivers\ZipDriver())
    {
        $this->driver = $driver;
    }

    /**
     * Generate a document from the template (document)
     *
     * @param string|\Stringable $templateFile
     * @param mixed $data
     * @return \AnourValar\Office\Generated
     */
    public function generate(string|\Stringable $templateFile, mixed $data): Generated
    {
        // Handle with input data
        $data = $this->canonizeData($data);

        // Open the template
        $templateFormat = Format::tryFrom(mb_strtolower(pathinfo($templateFile, PATHINFO_EXTENSION))) ?? Format::Docx;
        $driver = $this->driver->load($templateFile, $templateFormat);

        // Handle
        $driver->replace($data);

        // Return
        return new Generated($driver);
    }

    /**
     * @param mixed $data
     * @return array
     */
    protected function canonizeData(mixed $data): array
    {
        $result = [];

        if (is_object($data) && method_exists($data, 'toArray')) {
            $data = $data->toArray();
        }

        foreach ($this->dot($data) as $key => $value) {
            $result["[$key]"] = $value;
        }

        return $result;
    }
}
