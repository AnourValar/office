<?php

namespace AnourValar\Office\Drivers;

class PhpSpreadsheetDriver implements TemplateInterface, MixInterface
{
    /**
     * @var string
     */
    protected const DATE_FORMAT = 'm/d/yyyy';

    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public readonly \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public readonly \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet;

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\LoadInterface::load()
     */
    public function load(string $file, \AnourValar\Office\Format $format): self
    {
        $this->spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($this->getFormat($format))->load($file);
        $this->sheet = $this->spreadsheet->getActiveSheet();

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SaveInterface::save()
     */
    public function save(string $file, \AnourValar\Office\Format $format): void
    {
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, $this->getFormat($format));

        if (method_exists($writer, 'writeAllSheets')) {
            $writer->writeAllSheets();
        }

        $writer->save($file);
    }

    /**
     * Apply value to a cell
     *
     * @param string $cell
     * @param mixed $value
     * @return self
     */
    public function setValue(string $cell, $value): self
    {
        if ($value instanceof \DateTimeInterface) {

            $this->sheet->setCellValue($cell, \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($value));
            $this->sheet->getStyle($cell)->getNumberFormat()->setFormatCode(static::DATE_FORMAT);

        } elseif (is_string($value) && preg_match('#^\=[A-Z][A-Z\.\d]#', $value)) {

            $this->sheet->setCellValue($cell, $value);

        } elseif (is_string($value) || is_null($value)) {

            $this->sheet->getCell($cell)->setValueExplicit((string) $value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);

        } else {

            if (is_double($value)) {
                if (abs($value) >= 1000) {
                    $this->setCellFormat($cell, '# ##0.00');
                } else {
                    $this->setCellFormat($cell, '0.00');
                }
            } elseif (is_integer($value)) {
                if (abs($value) >= 1000) {
                    $this->setCellFormat($cell, '# ##0');
                } else {
                    $this->setCellFormat($cell, '0');
                }
            }

            $this->sheet->getCell($cell)->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);

        }

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::setValues()
     */
    public function setValues(array $data): self
    {
        foreach ($data as $row => $columns) {
            foreach ($columns as $column => $value) {
                $this->setValue($column.$row, $value);
            }
        }

        return $this;
    }

    /**
     * Get cell' value
     *
     * @param string $cell
     * @return mixed
     */
    public function getValue(string $cell)
    {
        return $this->sheet->getCell($cell)->getValue();
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::getValues()
     */
    public function getValues(?string $ceilRange): array
    {
        if (! $ceilRange) {
            $ceilRange = sprintf('A1:%s%s', $this->sheet->getHighestColumn(), $this->sheet->getHighestRow());
        }

        return $this->sheet->rangeToArray(
            $ceilRange, // The worksheet range that we want to retrieve
            null,       // Value that should be returned for empty cells
            false,      // Should formulas be calculated (the equivalent of getCalculatedValue() for each cell)
            false,      // Should values be formatted (the equivalent of getFormattedValue() for each cell)
            true        // Should the array be indexed by cell row and cell column
        );
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::getMergeCells()
     */
    public function getMergeCells(): array
    {
        return array_values( $this->sheet->getMergeCells() );
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::mergeCells()
     */
    public function mergeCells(string $ceilRange): self
    {
        $this->sheet->mergeCells($ceilRange);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::copyStyle()
     */
    public function copyStyle(string $cellFrom, string $rangeTo): self
    {
        $this->sheet->duplicateStyle($this->sheet->getStyle($cellFrom), $rangeTo);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::addRow()
     */
    public function addRow(int $rowBefore, int $qty = 1): self
    {
        $this->sheet->insertNewRowBefore($rowBefore, $qty);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::deleteRow()
     */
    public function deleteRow(int $row, int $qty = 1): self
    {
        $this->sheet->removeRow($row, $qty);

        return $this;
    }

    /**
     * Add a column
     *
     * @param string $columnBefore
     * @param int $qty
     * @return self
     */
    public function addColumn(string $columnBefore, int $qty = 1): self
    {
        $this->sheet->insertNewColumnBefore($columnBefore, $qty);

        return $this;
    }

    /**
     * Set auto-width for a column
     *
     * @param string $column
     * @return self
     */
    public function autoWidth(string $column): self
    {
        $this->sheet->getColumnDimension($column)->setAutoSize(true);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\TemplateInterface::setWidth()
     */
    public function setWidth(string $column, mixed $width = null): self
    {
        if (is_null($width)) {
            $width = mb_strlen($width) * 1.2 + ( (mb_strlen((int) $width) < 10) ? 4 : 0 );
        }

        if (is_string($width)) {
            $width = $this->sheet->getColumnDimension($width)->getWidth();
        }

        $this->sheet->getColumnDimension($column)->setWidth($width);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MixInterface::setSheetTitle()
     */
    public function setSheetTitle(string $title): self
    {
        $this->sheet->setTitle($title);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MixInterface::getSheetTitle()
     */
    public function getSheetTitle(): string
    {
        return $this->sheet->getTitle();
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MixInterface::mergeDriver()
     */
    public function mergeDriver(\AnourValar\Office\Drivers\MixInterface $driver): self
    {
        $this->spreadsheet->addExternalSheet($driver->sheet);

        return $this;
    }

    /**
     * Place an image
     *
     * @param string $filename
     * @param string $cell
     * @param array $options
     * @return self
     */
    public function insertImage(string $filename, string $cell, array $options = []): self
    {
        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();

        $drawing->setPath($filename); // put your path and image here
        $drawing->setCoordinates($cell);

        if (isset($options['name'])) {
            $drawing->setName($options['name']);
        }

        if (isset($options['offset_x'])) {
            $drawing->setOffsetX($options['offset_x']);
        }

        if (isset($options['offset_y'])) {
            $drawing->setOffsetY($options['offset_y']);
        }

        if (isset($options['rotation'])) {
            $drawing->setRotation($options['rotation']);
        }

        if (isset($options['width']) && isset($options['height'])) {
            $drawing
                ->setResizeProportional(false)
                ->setWidth($options['width'])
                ->setHeight($options['height']);
        } elseif (isset($options['width'])) {
            $drawing->setWidth($options['width']);
        } elseif (isset($options['height'])) {
            $drawing->setHeight($options['height']);
        }

        $drawing->setWorksheet($this->sheet);

        return $this;
    }

    /**
     * @param \AnourValar\Office\Format $format
     * @return string
     */
    protected function getFormat(\AnourValar\Office\Format $format): string
    {
        return match($format) {
            \AnourValar\Office\Format::Xlsx => 'Xlsx',
            \AnourValar\Office\Format::Pdf => 'Mpdf',
            \AnourValar\Office\Format::Html => 'Html',
            \AnourValar\Office\Format::Ods => 'Ods',
        };
    }

    /**
     * Установка формата ячейки
     *
     * @param string $cell
     * @param string $format
     * @return void
     */
    private function setCellFormat(string $cell, string $format): void
    {
        $this
            ->sheet
            ->getStyle($cell)
            ->getNumberFormat()
            ->setFormatCode($format);
    }
}
