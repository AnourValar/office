<?php

namespace AnourValar\Office\Drivers;

use PhpOffice\PhpSpreadsheet\Style\Border;

class PhpSpreadsheetDriver implements TemplateInterface, GridInterface, MixInterface
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
     * @see \AnourValar\Office\Drivers\GridInterface::create()
     */
    public function create(): self
    {
        $this->spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();

        return $this;
    }

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
        $this->sheet->setSelectedCells('A1');

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
     * @param string $format
     * @return self
     */
    public function setValue(string $cell, $value, string $format = null): self
    {
        if ($value instanceof \DateTimeInterface) {

            $this->sheet->setCellValue($cell, \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($value));
            $this->setCellFormat($cell, ($format ?? static::DATE_FORMAT));

        } elseif (is_string($value) || is_null($value)) {

            if (is_numeric($value)) {
                $this->sheet->getCell($cell)->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
            } else {
                $this->sheet->setCellValue($cell, $value);
            }

        } else {

            if (is_double($value)) {
                $this->setCellFormat($cell, ($format ?? '#,##0.00'));
            } elseif (is_integer($value)) {
                $this->setCellFormat($cell, ($format ?? '#,##0'));
            }

            $this->sheet->getCell($cell)->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_NUMERIC);

        }

        return $this;
    }

    /**
     * Set data format of cell (column, range)
     *
     * @param string $range
     * @param string $format
     * @return self
     */
    public function setCellFormat(string $range, string $format): self
    {
        $this->sheet->getStyle($range)->getNumberFormat()->setFormatCode($format);

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
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\GridInterface::setGrid()
     */
    public function setGrid(iterable $data): self
    {
        $row = 0;
        foreach ($data as $values) {
            $row++;
            $column = 'A';

            foreach ($values as $value) {
                if ($value !== '' && $value !== null) {
                    $this->setValue($column.$row, $value);
                }

                $column++;
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
    public function setWidth(string $column, int|string $width): self
    {
        if (is_string($width)) {
            $width = $this->sheet->getColumnDimension($width)->getWidth();
        }

        $this->sheet->getColumnDimension($column)->setWidth($width);

        return $this;
    }

    /**
     * Set fixed height for a row
     *
     * @param string $row
     * @param int|string $height
     * @return self
     */
    public function setHeight(string $row, int|string $height): self
    {
        if (mb_substr($height, 0, 1) == '#') {
            $height = $this->sheet->getRowDimension(mb_substr($height, 1))->getRowHeight();
        }

        $this->sheet->getRowDimension($row)->setRowHeight($height);

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
     * Set custom style for the range of cells
     *
     * @param string $range
     * @param array $style
     * @return self
     */
    public function setStyle(string $range, array $style): self
    {
        if (isset($style['bold'])) {
            $this->sheet->getStyle($range)->getFont()->setBold($style['bold']);
        }

        if (isset($style['italic'])) {
            $this->sheet->getStyle($range)->getFont()->setItalic($style['italic']);
        }

        if (isset($style['size'])) {
            $this->sheet->getStyle($range)->getFont()->setSize($style['size']);
        }

        if (isset($style['underline'])) {
            $this->sheet->getStyle($range)->getFont()->setUnderline($style['underline']);
        }

        if (isset($style['color'])) {
            $this->sheet->getStyle($range)->getFont()->getColor()->setRGB($style['color']);
        }

        if (isset($style['background_color'])) {
            $this
                ->sheet
                ->getStyle($range)
                ->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getStartColor()
                ->setRGB($style['background_color']);
        }

        if (isset($style['borders'])) {
            $this
                ->sheet
                ->getStyle($range)
                ->getBorders()
                ->getAllBorders()
                ->setBorderStyle($style['borders'] ? Border::BORDER_THIN : Border::BORDER_NONE);
        }

        if (isset($style['borders_outline'])) {
            $this
                ->sheet
                ->getStyle($range)
                ->getBorders()
                ->getOutline()
                ->setBorderStyle($style['borders_outline'] ? Border::BORDER_THIN : Border::BORDER_NONE);
        }

        if (isset($style['align'])) {
            $align = match($style['align'])
            {
                \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT => 'left',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER => 'center',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT => 'right',
            };

            $this
                ->sheet
                ->getStyle($range)
                ->getAlignment()->setHorizontal($align);
        }

        if (isset($style['valign'])) {
            $valign = match($style['valign'])
            {
                \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP => 'top',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER => 'center',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM => 'bottom',
            };

            $this
                ->sheet
                ->getStyle($range)
                ->getAlignment()->setVertical($valign);
        }

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
}
