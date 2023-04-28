<?php

namespace AnourValar\Office\Drivers;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;

class PhpSpreadsheetDriver implements SheetsInterface, GridInterface, MixInterface
{
    use \AnourValar\Office\Traits\Parser;

    /**
     * @var string
     */
    protected const FORMAT_DATE = 'm/d/yyyy';

    /**
     * @var string
     */
    protected const FORMAT_DOUBLE = '#,##0.00';

    /**
     * @var string
     */
    protected const FORMAT_INT = '#,##0';

    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    public readonly \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet;

    /**
     * @var int
     */
    protected int $sourceActiveSheetIndex;

    /**
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function sheet(): \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
    {
        return $this->spreadsheet->getActiveSheet();
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\GridInterface::create()
     */
    public function create(): self
    {
        $instance = new static;
        $instance->spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $instance->sourceActiveSheetIndex = 0;

        $this->readConfiguration($instance);
        return $instance;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\LoadInterface::load()
     */
    public function load(string $file, \AnourValar\Office\Format $format): self
    {
        $instance = new static;
        $instance->spreadsheet = IOFactory::createReader($instance->getFormat($format))->load($file);
        $instance->sourceActiveSheetIndex = $instance->spreadsheet->getActiveSheetIndex();

        $this->readConfiguration($instance);
        return $instance;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SaveInterface::save()
     */
    public function save(string $file, \AnourValar\Office\Format $format): void
    {
        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($this->spreadsheet, $this->getFormat($format));
        $this->writeConfiguration($writer);

        $count = $this->spreadsheet->getSheetCount();
        for ($i = 0; $i < $count; $i++) {
            $this->spreadsheet->getSheet($i)->setSelectedCells('A1');
        }
        $this->spreadsheet->setActiveSheetIndex($this->sourceActiveSheetIndex);

        if (method_exists($writer, 'writeAllSheets')) {
            $writer->writeAllSheets();
        }

        $writer->save($file);
    }

    /**
     * Clean up
     *
     * @return void
     */
    public function __destruct()
    {
        if (isset($this->spreadsheet)) {
            $this->spreadsheet->disconnectWorksheets();
            gc_collect_cycles();
        }
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MultiSheetInterface::setSheet()
     */
    public function setSheet(int $index): self
    {
        $this->spreadsheet->setActiveSheetIndex($index);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MultiSheetInterface::getSheetCount()
     */
    public function getSheetCount(): int
    {
        return $this->spreadsheet->getSheetCount();
    }

    /**
     * Apply value to a cell
     *
     * @param string $cell
     * @param mixed $value
     * @param bool $autoCellFormat
     * @return self
     */
    public function setValue(string $cell, $value, bool $autoCellFormat = true): self
    {
        if ($value instanceof \DateTimeInterface) {

            $this->sheet()->setCellValue($cell, \PhpOffice\PhpSpreadsheet\Shared\Date::PHPToExcel($value));
            if ($autoCellFormat) {
                $this->setCellFormat($cell, static::FORMAT_DATE);
            }

        } elseif (is_string($value) || is_null($value)) {

            if (is_numeric($value)) {
                $this->sheet()->getCell($cell)->setValueExplicit($value, DataType::TYPE_STRING);
            } else {
                $this->sheet()->setCellValue($cell, $value);
            }

        } else {

            if ($autoCellFormat && is_double($value)) {
                $this->setCellFormat($cell, static::FORMAT_DOUBLE);
            } elseif ($autoCellFormat && is_integer($value)) {
                $this->setCellFormat($cell, static::FORMAT_INT);
            }

            $this->sheet()->getCell($cell)->setValueExplicit($value, DataType::TYPE_NUMERIC);

        }

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::setValues()
     */
    public function setValues(array $data, bool $autoCellFormat = true): self
    {
        foreach ($data as $row => $columns) {
            foreach ($columns as $column => $value) {
                $this->setValue($column.$row, $value, $autoCellFormat);
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
        return $this->sheet()->getCell($cell)->getValue();
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::getValues()
     */
    public function getValues(?string $ceilRange): array
    {
        if (! $ceilRange) {
            $ceilRange = sprintf('A1:%s%s', $this->sheet()->getHighestColumn(), $this->sheet()->getHighestRow());
        }

        return $this->sheet()->rangeToArray(
            $ceilRange, // The worksheet range that we want to retrieve
            null,       // Value that should be returned for empty cells
            false,      // Should formulas be calculated (the equivalent of getCalculatedValue() for each cell)
            false,      // Should values be formatted (the equivalent of getFormattedValue() for each cell)
            true        // Should the array be indexed by cell row and cell column
        );
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::getMergeCells()
     */
    public function getMergeCells(): array
    {
        return array_values( $this->sheet()->getMergeCells() );
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::mergeCells()
     */
    public function mergeCells(string $ceilRange): self
    {
        $this->sheet()->mergeCells($ceilRange);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::copyStyle()
     */
    public function copyStyle(string $cellFrom, string $rangeTo): self
    {
        $this->sheet()->duplicateStyle($this->sheet()->getStyle($cellFrom), $rangeTo);

        if ($conditionalStyle = $this->sheet()->getConditionalStyles($cellFrom)) {
            $this->sheet()->duplicateConditionalStyle($conditionalStyle, $rangeTo);
        }

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::copyCellFormat()
     */
    public function copyCellFormat(string $cellFrom, string $rangeTo): self
    {
        $this->setCellFormat($rangeTo, $this->sheet()->getStyle($cellFrom)->getNumberFormat()->getFormatCode());

        return $this;
    }

    /**
     * Set cell (data) format
     *
     * @param string $range
     * @param string $format
     * @return self
     */
    public function setCellFormat(string $range, string $format): self
    {
        $this->sheet()->getStyle($range)->getNumberFormat()->setFormatCode($format);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::addRow()
     */
    public function addRow(int $rowBefore, int $qty = 1): self
    {
        $this->sheet()->insertNewRowBefore($rowBefore, $qty);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::deleteRow()
     */
    public function deleteRow(int $row, int $qty = 1): self
    {
        foreach ($this->getMergeCells() as $merge) {
            preg_match('#(\d+)#', $merge, $details);
            if ($details[1] == $row) {
                $this->sheet()->unmergeCells($merge);
            }
        }

        $this->sheet()->removeRow($row, $qty);

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
        $this->sheet()->insertNewColumnBefore($columnBefore, $qty);

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
        $this->sheet()->getColumnDimension($column)->setAutoSize(true);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SheetsInterface::copyWidth()
     */
    public function copyWidth(string $columnFrom, string $columnTo): self
    {
        $width = $this->sheet()->getColumnDimension($columnFrom)->getWidth();
        $this->setWidth($columnTo, $width);

        return $this;
    }

    /**
     * Set fixed width for a column
     *
     * @param string $column
     * @param int|float $width
     * @return self
     */
    public function setWidth(string $column, int|float $width): self
    {
        $this->sheet()->getColumnDimension($column)->setWidth($width);

        return $this;
    }

    /**
     * Copy row's height
     *
     * @param int $rowFrom
     * @param int $rowTo
     * @return self
     */
    public function copyHeight(int $rowFrom, int $rowTo): self
    {
        $height = $this->sheet()->getRowDimension($rowFrom)->getRowHeight();
        $this->setHeight($rowTo, $height);

        return $this;
    }

    /**
     * Set fixed height for a row
     *
     * @param string $row
     * @param int|float $height
     * @return self
     */
    public function setHeight(string $row, int|float $height): self
    {
        $this->sheet()->getRowDimension($row)->setRowHeight($height);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MixInterface::setSheetTitle()
     */
    public function setSheetTitle(string $title): self
    {
        $this->sheet()->setTitle($title);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MixInterface::getSheetTitle()
     */
    public function getSheetTitle(): string
    {
        return $this->sheet()->getTitle();
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\MixInterface::mergeDriver()
     */
    public function mergeDriver(\AnourValar\Office\Drivers\MixInterface $driver): self
    {
        $index = $driver->spreadsheet->getActiveSheetIndex();
        $this->spreadsheet->addExternalSheet($driver->sheet());
        $driver->spreadsheet->createSheet($index);

        return $this;
    }

    /**
     * Apply cell`s style without format
     *
     * @param string $cellFrom
     * @param string $rangeTo
     * @param bool $copyAlignment
     * @return self
     */
    public function copyStyleWithoutFormat(string $cellFrom, string $rangeTo, bool $copyAlignment = false): self
    {
        $style = $this->sheet()->getStyle($cellFrom)->exportArray();
        if (! $copyAlignment) {
            unset($style['alignment'], $style['numberFormat'], $style['protection']);
        } else {
            unset($style['numberFormat'], $style['protection']);
        }

        // @TODO: fixed ?
        if (
            ! isset($style['borders']['allBorders'])
            && isset($style['borders']['bottom'], $style['borders']['top'])
            && isset($style['borders']['left'], $style['borders']['right'])
            && $style['borders']['bottom'] == $style['borders']['top']
            && $style['borders']['bottom'] == $style['borders']['left']
            && $style['borders']['bottom'] == $style['borders']['right']
        ) {
            $style['borders']['allBorders'] = $style['borders']['bottom'];

            unset($style['borders']['bottom'], $style['borders']['top']);
            unset($style['borders']['left'], $style['borders']['right']);
        }

        $this->sheet()->getStyle($rangeTo)->applyFromArray($style);

        return $this;
    }

    /**
     * Find a cell with the value
     *
     * @param mixed $value
     * @param bool $strict
     * @return array|null
     */
    public function findCell($value, bool $strict = false): ?array
    {
        foreach ($this->getValues(null) as $row => $rowData) {
            foreach ($rowData as $column => $columnData) {
                if (($strict && $columnData === $value) || (! $strict && $columnData == $value)) {
                    return [$column, $row];
                }
            }
        }

        return null;
    }

    /**
     * Duplicate rows (with style, value) by range
     *
     * @param string $ceilRange
     * @param callable $value
     * @param int $indentRows
     * @param bool $addRows
     * @return self
     */
    public function duplicateRows(string $ceilRange, callable $value, int $indentRows = 0, bool $addRows = true): self
    {
        $range = explode(':', $ceilRange);
        $range[0] = preg_split('#([A-Z]+)([\d]+)#S', $range[0], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
        $range[1] = preg_split('#([A-Z]+)([\d]+)#S', $range[1], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);

        $shift = $range[1][1] - $range[0][1] + 1 + $indentRows;
        $mergeCells = $this->getMergeCells();
        $values = $this->getValues($ceilRange);

        // Rows
        if ($addRows) {
            $this->addRow($range[1][1] + 1, $shift + $indentRows);
        }

        // Merge
        foreach ($mergeCells as $item) {
            $item = explode(':', $item);
            $item[0] = preg_split('#([A-Z]+)([\d]+)#S', $item[0], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
            $item[1] = preg_split('#([A-Z]+)([\d]+)#S', $item[1], -1, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);

            if (
                $item[0][1] >= $range[0][1] && $item[0][1] <= $range[1][1] // rows
                && $item[1][1] >= $range[0][1] && $item[1][1] <= $range[1][1]
                && $this->isColumnGE($item[0][0], $range[0][0]) && $this->isColumnLE($item[0][0], $range[1][0]) // columns
                && $this->isColumnGE($item[1][0], $range[0][0]) && $this->isColumnLE($item[1][0], $range[1][0])
            ) {
                $this->mergeCells($item[0][0].($item[0][1]+$shift) . ':' . $item[1][0].($item[1][1]+$shift));
            }
        }

        $curr = $range[0][1];
        while ($curr <= $range[1][1]) {
            // Height
            $this->copyHeight($curr, $curr + $shift);

            // Style, CellFormat, Value
            $column = $range[0][0];
            while ($this->isColumnLE($column, $range[1][0])) {
                $this->copyStyle($column . $curr, $column . ($curr + $shift));
                $this->copyCellFormat($column . $curr, $column . ($curr + $shift));

                if (isset($values[$curr][$column])) {
                    $this->setValue($column . ($curr + $shift), $value($values[$curr][$column], $column, $curr), false);
                }

                $column++;
            }

            $curr++;
        }

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
            $this->sheet()->getStyle($range)->getFont()->setBold($style['bold']);
        }

        if (isset($style['italic'])) {
            $this->sheet()->getStyle($range)->getFont()->setItalic($style['italic']);
        }

        if (isset($style['size'])) {
            $this->sheet()->getStyle($range)->getFont()->setSize($style['size']);
        }

        if (isset($style['underline'])) {
            $this->sheet()->getStyle($range)->getFont()->setUnderline($style['underline']);
        }

        if (isset($style['color'])) {
            $this->sheet()->getStyle($range)->getFont()->getColor()->setRGB($style['color']);
        }

        if (isset($style['background_color'])) {
            $this
                ->sheet()
                ->getStyle($range)
                ->getFill()
                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                ->getStartColor()
                ->setRGB($style['background_color']);
        }

        if (isset($style['borders'])) {
            $this
                ->sheet()
                ->getStyle($range)
                ->getBorders()
                ->getAllBorders()
                ->setBorderStyle($style['borders'] ? Border::BORDER_THIN : Border::BORDER_NONE);
        }

        if (isset($style['borders_outline'])) {
            $this
                ->sheet()
                ->getStyle($range)
                ->getBorders()
                ->getOutline()
                ->setBorderStyle($style['borders_outline'] ? Border::BORDER_THIN : Border::BORDER_NONE);
        }

        if (isset($style['align'])) {
            $align = match($style['align']) {
                \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT => 'left',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER => 'center',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT => 'right',
            };

            $this
                ->sheet()
                ->getStyle($range)
                ->getAlignment()->setHorizontal($align);
        }

        if (isset($style['valign'])) {
            $valign = match($style['valign']) {
                \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_TOP => 'top',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER => 'center',
                \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_BOTTOM => 'bottom',
            };

            $this
                ->sheet()
                ->getStyle($range)
                ->getAlignment()->setVertical($valign);
        }

        if (isset($style['wrap'])) {
            $this
                ->sheet()
                ->getStyle($range)
                ->getAlignment()->setWrapText($style['wrap']);
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

        if (isset($options['base64'])) {
            $filename = 'data:image/' . $options['base64'] . ';base64,' . base64_encode(file_get_contents($filename));
        }

        $drawing->setPath($filename);
        $drawing->setCoordinates($cell);

        if (isset($options['coordinates2'])) {
            $drawing->setCoordinates2($options['coordinates2']);
        }

        if (isset($options['name'])) {
            $drawing->setName($options['name']);
        }

        if (isset($options['offset_x'])) {
            $drawing->setOffsetX($options['offset_x']);
        }
        if (isset($options['offset_x2'])) {
            $drawing->setOffsetX2($options['offset_x2']);
        }

        if (isset($options['offset_y'])) {
            $drawing->setOffsetY($options['offset_y']);
        }
        if (isset($options['offset_y2'])) {
            $drawing->setOffsetY2($options['offset_y2']);
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

        $drawing->setWorksheet($this->sheet());

        return $this;
    }

    /**
     * "Reader" configuration
     *
     * @param \AnourValar\Office\Drivers\PhpSpreadsheetDriver $instance
     * @return void
     */
    protected function readConfiguration(PhpSpreadsheetDriver $instance): void
    {
        //
    }

    /**
     * "Writer" configuration
     *
     * @param \PhpOffice\PhpSpreadsheet\Writer\IWriter $writer
     * @return void
     */
    protected function writeConfiguration(\PhpOffice\PhpSpreadsheet\Writer\IWriter $writer): void
    {
        //
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
