<?php

namespace AnourValar\Office\Drivers;

class ZipDriver implements DocumentInterface, GridInterface
{
    use \Anourvalar\Office\Traits\Parser;
    use \AnourValar\Office\Traits\XFormat;

    /**
     * @var \AnourValar\Office\Format
     */
    protected readonly \AnourValar\Office\Format $format;

    /**
     * @var array
     */
    protected array $fileSystem;

    /**
     * @var array
     */
    protected array $gridOptions = [];

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\GridInterface::create()
     */
    public function create(): self
    {
        return $this->load(__DIR__ . '/../resources/grid.xlsx', \AnourValar\Office\Format::Xlsx);
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\LoadInterface::load()
     */
    public function load(string $file, \AnourValar\Office\Format $format): self
    {
        if (! in_array($format, [\AnourValar\Office\Format::Docx, \AnourValar\Office\Format::Xlsx])) {
            throw new \LogicException('Driver only supports Docx, Xlsx formats.');
        }

        $instance = new static;
        $fileSystem = [];

        $zipArchive = new \ZipArchive;
        $zipArchive->open($file);
        try {
            $count = $zipArchive->numFiles;

            for ($i = 0; $i < $count; $i++) {
                $filename = $zipArchive->getNameIndex($i);
                $content = $zipArchive->getFromName($filename);

                $fileSystem[$filename] = $content;
            }
        } catch (\Throwable $e) {
            $zipArchive->close();
            throw $e;
        }
        $zipArchive->close();

        $instance->fileSystem = $fileSystem;
        $instance->format = $format;
        return $instance;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\SaveInterface::save()
     */
    public function save(string $file, \AnourValar\Office\Format $format): void
    {
        if ($format != $this->format) {
            throw new \LogicException('Driver only supports saving in the same format.');
        }

        $zipStream = new \ZipStream\ZipStream();
        ob_start();

        try {
            foreach ($this->fileSystem as $filename => $content) {
                $zipStream->addFile($filename, $content);
             }
        } catch (\Throwable $e) {
            $zipStream->finish();
            ob_get_clean();
            throw $e;
        }

        $zipStream->finish();
        file_put_contents($file, ob_get_clean());
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\DocumentInterface::replace()
     */
    public function replace(array $data): self
    {
        foreach ($data as &$value) {
            $value = $this->escape($value);
        }
        unset($value);

        foreach ($this->fileSystem as $filename => &$content) {
            if ($content && mb_strtolower(pathinfo($filename, PATHINFO_EXTENSION)) == 'xml') {
                $content = $this->handleReplace($content, $data);
            }
        }
        unset($content);

        return $this;
    }

    /**
     * {@inheritDoc}
     * @see \AnourValar\Office\Drivers\GridInterface::setGrid()
     */
    public function setGrid(iterable $data): self
    {
        $sheet = ''; // matrix
        $cols = ''; // columns

        $buckets = []; // text (sharedStrings)
        $bucketsIndex = 0;

        $row = 0; // rows count
        $columnsCount = 0; // columns count
        $firstColumn = 'A'; // columns shift

        // Styles
        $styles = $this->loadGridStyles();

        // Head (titles)
        while ($data->valid()) {
            $row++;
            $titles = $data->current();
            $data->next();
            $columnsCount = count($titles);

            if (! $titles) {
                continue;
            }

            $sheet .= '<row r="'.$row.'" spans="1:'.$columnsCount.'" ht="25" customHeight="1" x14ac:dyDescent="0.3">';
            $column = 'A';
            foreach ($titles as $value) {
                $value = (string) $value;
                if ($value === null || $value === '') {
                    $firstColumn++;
                    $column++;
                    continue;
                }

                $value = $this->escape($value);
                $curr = $buckets[$value] ?? null;
                if ($curr === null) {
                    $curr = $bucketsIndex;
                    $buckets[$value] = $curr;

                    $bucketsIndex++;
                }

                $sheet .= '<c r="'.$column.$row.'" t="s" s="'.$styles['header'].'"><v>'.$curr.'</v></c>';
                $column++;
            }
            $sheet .= '</row>';

            break;
        }

        // Custom styles
        foreach (($this->gridOptions['style'] ?? []) as $column => $alias) {
            if (isset($styles[$alias])) {
                $styles[$column] = $styles[$alias];
            }
        }

        // Body (data)
        while ($data->valid()) {
            $values = $data->current();
            $data->next();
            $row++;

            $sheet .= '<row r="'.$row.'" spans="1:'.$columnsCount.'" x14ac:dyDescent="0.3">';
            $column = 'A';
            foreach ($values as $value) {
                if ($value instanceof \Stringable && ! $value instanceof \DateTimeInterface) {
                    $value = (string) $value;
                }

                if ($value === null || $value === '') {

                    if ($this->isColumnGE($column, $firstColumn)) {
                        $sheet .= '<c r="'.$column.$row.'" s="'.($styles[$column] ?? $styles['string']).'"/>';
                    }

                } elseif (is_string($value)) {

                    $value = $this->escape($value);
                    $curr = $buckets[$value] ?? null;
                    if ($curr === null) {
                        $curr = $bucketsIndex;
                        $buckets[$value] = $curr;

                        $bucketsIndex++;
                    }

                    $style = ($styles[$column] ?? $styles['string']);
                    $sheet .= '<c r="'.$column.$row.'" t="s" s="'.$style.'"><v>'.$curr.'</v></c>';

                } elseif (is_float($value)) {

                    $style = ($styles[$column] ?? $styles['float']);
                    $sheet .= '<c r="'.$column.$row.'" s="'.$style.'"><v>'.$value.'</v></c>';

                } elseif (is_integer($value)) {

                    $style = ($styles[$column] ?? $styles['integer']);
                    $sheet .= '<c r="'.$column.$row.'" s="'.$style.'"><v>'.$value.'</v></c>';

                } elseif ($value instanceof \DateTimeInterface) {

                    $style = ($styles[$column] ?? $styles['date']);
                    $sheet .= '<c r="'.$column.$row.'" s="'.$style.'"><v>'.$this->excelDate($value).'</v></c>';

                } else {

                    throw new \RuntimeException('Unsupported type of value.');

                }

                $column++;
            }
            $sheet .= '</row>';
        }

        // Columns
        $column = 'A';
        for ($index = 1; $index <= $columnsCount; $index++) {
            if ($this->isColumnGE($column, $firstColumn)) {
                $width = ($this->gridOptions['width'][$column] ?? 20);
                $cols .= '<col min="'.$index.'" max="'.$index.'" width="'.$width.'" customWidth="1"/>';
            }
            $column++;
        }

        // Save buckets
        $this->saveGridSharedStrings($buckets);

        // Save columns & matrix
        $this->saveGridWorksheet($cols, $sheet, $row, $column);

        // Etc
        $this->saveGridEtc();

        return $this;
    }

    /**
     * @param string $content
     * @param array $data
     * @return string
     */
    protected function handleReplace(string $content, array &$data): string
    {
        foreach ($data as $from => $to) {
            $pattern = mb_str_split($from);
            foreach ($pattern as &$patternItem) {
                $patternItem = preg_quote($patternItem);
                $patternItem .= '(\<[^\[]*)?';
            }
            unset($patternItem);
            $pattern = implode('', $pattern);

            $content = preg_replace_callback("#$pattern#Uu", function ($patterns) use ($from, $to) {
                if (strip_tags($patterns[0]) == $from) {
                    return $to;
                }

                return $patterns[0];
            }, $content);
        }

        return $content;
    }


    /**
     * Set styles map for the grid
     *
     * @param string $column
     * @param string $style
     * @return self
     */
    public function setStyle(string $column, string $style): self
    {
        $this->gridOptions['style'][$column] = $style;

        return $this;
    }

    /**
     * Set column's width for the grid
     *
     * @param string $column
     * @param int $width
     * @return self
     */
    public function setWidth(string $column, int $width): self
    {
        $this->gridOptions['width'][$column] = $width;

        return $this;
    }

    /**
     * Set sheet title for the grid
     *
     * @param string $title
     * @return self
     */
    public function setSheetTitle(string $title): self
    {
        $this->fileSystem['xl/workbook.xml'] = preg_replace(
            '#\<sheet name="(.+)" sheetId\="(.+)" r\:id\="(.+)"\/\>#uU',
            '<sheet name="'.$this->escape($title).'" sheetId="$2" r:id="$3"/>',
            $this->fileSystem['xl/workbook.xml']
        );

        return $this;
    }

    /**
     * @return array
     */
    protected function loadGridStyles(): array
    {
        $styles = [];

        // Bucket
        preg_match_all('#<si><t>(.*)</t></si>#uU', $this->fileSystem['xl/sharedStrings.xml'], $buckets);
        $buckets = $buckets[1];

        // Matrix
        preg_match_all(
            '#<c r="[A-Z\d]+" s="(\d+)" t="s"><v>(\d+)</v></c>#uU',
            $this->fileSystem['xl/worksheets/sheet1.xml'],
            $matrix
        );

        // Parse the map
        foreach ($matrix[2] as $key => $value) {
            if (isset($buckets[$value])) {
                $styles[mb_substr($buckets[$value], 1, -1)] = $matrix[1][$key];
            }
        }

        // Presets
        return array_merge(['header' => 1, 'string' => 1, 'float' => 1, 'integer' => 1, 'date' => 1], $styles);
    }

    /**
     * @param array $buckets
     * @return void
     */
    protected function saveGridSharedStrings(array &$buckets): void
    {
        $sst = '';
        foreach (array_keys($buckets) as $word) {
            $sst .= '<si><t>'.$word.'</t></si>';
        }
        $count = count($buckets);

        $this->fileSystem['xl/sharedStrings.xml'] = <<<HERE
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="$count" uniqueCount="$count">
        $sst
        </sst>
        HERE;
    }

    /**
     * @param string $cols
     * @param string $sheet
     * @param int $lastRow
     * @param string $lastColumn
     * @return void
     */
    protected function saveGridWorksheet(string &$cols, string &$sheet, int $lastRow, string $lastColumn): void
    {
        $worksheet = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            . 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            . 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
            . 'mc:Ignorable="x14ac xr xr2 xr3" '
            . 'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" '
            . 'xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" '
            . 'xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" '
            . 'xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" '
            . 'xr:uid="{9CD2C5FF-272A-44F0-88FE-0A0D7B1A3B22}"';

        if ($cols) {
            $cols = "<cols>$cols</cols>";
        }

        if ($sheet) {
            $sheet = "<sheetData>$sheet</sheetData>";
        } else {
            $sheet = '<sheetData/>';
        }

        if (! $lastRow) {
            $lastRow = 1;
        }

        $this->fileSystem['xl/worksheets/sheet1.xml'] = <<<HERE
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet $worksheet>
        <dimension ref="A1:{$lastColumn}{$lastRow}"/>
        <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView>
        </sheetViews>
        <sheetFormatPr defaultRowHeight="14.4" x14ac:dyDescent="0.3"/>
        $cols
        $sheet
        <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        <pageSetup paperSize="9" orientation="portrait" horizontalDpi="1200" verticalDpi="1200" r:id="rId1"/>
        </worksheet>
        HERE;
    }

    /**
     * @return void
     */
    protected function saveGridEtc(): void
    {
        // Created at timestamp
        $this->fileSystem['docProps/core.xml'] = preg_replace(
            '#(\<dcterms\:created xsi\:type\="[^"]+">)(.*?)(\<\/dcterms\:created\>)#',
            '${1}' . date('Y-m-d\TH:i:s\Z') . '$3',
            $this->fileSystem['docProps/core.xml']
        );

        // Modified at timestamp
        $this->fileSystem['docProps/core.xml'] = preg_replace(
            '#(\<dcterms\:modified xsi\:type\="[^"]+">)(.*?)(\<\/dcterms\:modified\>)#',
            '${1}' . date('Y-m-d\TH:i:s\Z') . '$3',
            $this->fileSystem['docProps/core.xml']
        );
    }
}
