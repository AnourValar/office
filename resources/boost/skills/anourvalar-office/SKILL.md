---
name: anourvalar-office
description: Use when working with the anourvalar/office package to generate Excel/CSV/PDF/HTML/ODS documents from XLSX templates, fill DOCX (Word) templates, export tabular data (Grids), or merge generated spreadsheets. The package exposes plain service classes (SheetsService, DocumentService, GridService, Mixer) and a contract-based ExportService helper - it does NOT register Laravel facades.
---

# AnourValar Office

`anourvalar/office` is a framework-agnostic PHP 8.1+ library for generating documents from Excel/Word templates and exporting tabular data to Excel/CSV/PDF/HTML/ODS. It is commonly used in Laravel apps to stream reports, invoices, and contracts. The library has no service provider and registers no Laravel facades; consumers instantiate services directly (`new \AnourValar\Office\SheetsService()`) or wire them via Laravel's container.

## When to use

- Generating an `.xlsx`/`.ods`/`.pdf`/`.html`/`.csv` file from a pre-designed XLSX/ODS template with marker placeholders like `[foo]` or `[products.name]`.
- Filling a `.docx` template (Word) with values from a data array or Eloquent model.
- Exporting tables/reports (with headers + rows, optionally from a generator/Eloquent query) to spreadsheets.
- Merging multiple generated spreadsheet documents into a single workbook with multiple sheets.
- Streaming a download response from a Laravel controller via `response()->streamDownload(...)`.

Required optional dependencies (install only what you need):

- `phpoffice/phpspreadsheet ^3.10` - for XLSX/ODS/HTML/CSV read+write and PDF output (with `mpdf/mpdf ^8.1`).
- `maennchen/zipstream-php ^3.2` - for DOCX template filling and the fast `ZipDriver` XLSX grid writer.
- `mpdf/mpdf ^8.1` - for PDF export.

## Facades

This package does NOT ship Laravel facades extending `Illuminate\Support\Facades\Facade`. The `AnourValar\Office\Facades` namespace contains contracts/services that are instantiated or implemented manually:

- `AnourValar\Office\Facades\ExportService` - high-level helper that wraps `GridService` for Eloquent-style exports.
- `AnourValar\Office\Facades\ExportGridInterface` - contract you implement to describe an exportable grid.
- `AnourValar\Office\Facades\ExportGridQueryInterface` - extends `ExportGridInterface` for query-backed exports.

### `AnourValar\Office\Facades\ExportService`

Resolves nothing - construct it directly (`new ExportService()`) or bind it in the container.

Public methods:

- `grid(\Closure $dataGenerator, ExportGridInterface $grid, Format $format = Format::Xlsx, array $request = []): string` - runs the grid through the appropriate driver and returns the binary document. Picks `ZipDriver` for Xlsx and `PhpSpreadsheetDriver` for other formats.

Constants used as header-array flags to format whole columns:

- `ExportService::PERCENTAGE`
- `ExportService::DOUBLE_10`
- `ExportService::DATETIME`

```php
use AnourValar\Office\Facades\ExportService;
use AnourValar\Office\Format;

$binary = (new ExportService())->grid(
    fn () => yield from User::query()->cursor(),
    $myGrid,                  // implements ExportGridInterface
    Format::Xlsx,
    request()->all(),
);
```

### `AnourValar\Office\Facades\ExportGridInterface`

Implement this contract to drive `ExportService::grid()`:

- `sheetTitle(array $request): string` - title for the active sheet.
- `columns(array $request): array` - header definitions (each item is `['title' => ..., 'width' => ..., 'height' => ..., ExportService::PERCENTAGE => true, ...]`).
- `item($row, GridInterface $driver, int $rowNumber, array $request): array` - convert one source row into an array of cell values.
- `fileName(string $ext, array $request): string` - download filename including extension.

### `AnourValar\Office\Facades\ExportGridQueryInterface`

Extends `ExportGridInterface` and adds:

- `query(): \Illuminate\Database\Eloquent\Builder` - base Eloquent query to stream.

The interface's docblock shows the canonical Laravel streaming pattern (see Usage examples below).

## Services

### `AnourValar\Office\SheetsService`

Renders an XLSX/ODS template by replacing markers like `[scalar]`, `[group.product]`, and one/two-dimensional table markers, then saves to Xlsx/Pdf/Html/Ods.

Constructor: `__construct(SheetsInterface $driver = new PhpSpreadsheetDriver(), \AnourValar\Office\Sheets\Parser $parser = new \AnourValar\Office\Sheets\Parser())`.

Public methods:

- `generate(string|\Stringable $templateFile, mixed $data, bool $autoCellFormat = false): Generated` - renders and returns a `Generated` instance.
- `hookLoad(?\Closure $closure): self` - signature `function (SheetsInterface $driver, string $templateFile, Format $templateFormat)`; return a `SheetsInterface` (often `$driver->create()`) to build a template on the fly.
- `hookBefore(?\Closure $closure): self` - signature `function (SheetsInterface $driver, array &$data)`; mutate the driver/data before markers are written.
- `hookValue(?\Closure $closure): self` - signature `function (SheetsInterface $driver, string $column, int $row, mixed $value, int $sheetIndex)`; transform every cell value being written (return `null` to skip the cell).
- `hookAfter(?\Closure $closure): self` - signature `function (SheetsInterface $driver)`; post-processing after data is written.

### `AnourValar\Office\DocumentService`

Renders DOCX templates by literal marker replacement.

Constructor: `__construct(DocumentInterface $driver = new ZipDriver())`.

Public methods:

- `generate(string|\Stringable $templateFile, mixed $data): Generated` - flattens `$data` via dot notation, wraps keys as `[key]`, then replaces in the docx XML.

### `AnourValar\Office\GridService`

Builds a fresh spreadsheet from `[headers, rows]` data, optionally driven by hooks.

Constructor: `__construct(GridInterface $driver = new PhpSpreadsheetDriver())`. Pass `new ZipDriver()` for faster, low-memory XLSX output.

Public methods:

- `generate(array $headers, iterable|\Closure $data, string $leftTopCorner = 'A1'): Generated`.
- `hookLoad(?\Closure $closure)` - `function (GridInterface $driver): GridInterface`.
- `hookBefore(?\Closure $closure)` - `function (GridInterface $driver, array &$headers, iterable &$data, string $leftTopCorner)`.
- `hookHeader(?\Closure $closure)` - `function (GridInterface $driver, mixed $header, string|int $key, string $column, int $rowNumber): string` - return the header label to write.
- `hookRow(?\Closure $closure)` - `function (GridInterface $driver, mixed $row, string|int $key, int $rowNumber): array|null` - return an array of cell values (or `null` to skip).
- `hookAfter(?\Closure $closure)` - `function (GridInterface $driver, ?string $headersRange, ?string $dataRange, ?string $totalRange, array $columns)` - style ranges after all data is written.

### `AnourValar\Office\Generated`

Returned by every `generate(...)` call. Holds the underlying driver and saves output.

- `save(Format $format): string` - returns the binary document.
- `saveAs(string $filename, ?Format $format = null): ?int` - writes to disk; format inferred from extension when omitted.
- `hookSave(?\Closure $closure): self` - intercept saving with `function (SaveInterface $driver, Format $format)`.
- Public readonly property `$driver` - the underlying driver (e.g. `PhpSpreadsheetDriver` or `ZipDriver`).

### `AnourValar\Office\Mixer`

Invokable helper to union multiple `Generated` documents into one multi-sheet workbook.

- `__invoke(Generated ...$generated): Generated` - all drivers must be the same concrete class and implement `MixInterface`. Sheet titles are auto-suffixed `(1)`, `(2)`, ... on collision.

### `AnourValar\Office\Format`

`enum Format: string` with cases `Xlsx`, `Pdf`, `Html`, `Ods`, `Csv`, `Docx`. Methods:

- `fileExtension(): string`
- `contentType(): string` (MIME for HTTP responses)

### `AnourValar\Office\Buffer`

`implements \Stringable`. Wraps a binary string in a `tmpfile()` and stringifies to the temp filename. Useful when an API expects a path. The file is closed and removed on destruct.

### Drivers

All drivers live under `AnourValar\Office\Drivers`. The interfaces are the public contract consumers should rely on:

- `SheetsInterface` (extends `SaveInterface`, `LoadInterface`, `MultiSheetInterface`) - methods used in `hookBefore` / `hookValue`: `setValue(string $cell, $value, bool $autoCellFormat = true)`, `setValues(array $data, bool $autoCellFormat = true)`, `getValues(?string $ceilRange)`, `getMergeCells()`, `mergeCells(string $ceilRange)`, `copyStyle(...)`, `copyCellFormat(...)`, `addRow(int $rowBefore, int $qty = 1)`, `deleteRow(int $row, int $qty = 1)`, `copyWidth(string $columnFrom, string $columnTo)`.
- `GridInterface` (extends `SaveInterface`) - `create(): self`, `setGrid(iterable $data): self`.
- `DocumentInterface` (extends `SaveInterface`, `LoadInterface`) - `replace(array $data): self`.
- `MixInterface` (extends `MultiSheetInterface`) - `setSheetTitle`, `getSheetTitle`, `mergeDriver`.
- `MultiSheetInterface` - `setSheet(int $index)`, `getSheetCount()`.

Concrete drivers:

- `PhpSpreadsheetDriver` (default for `SheetsService` and `GridService`) - implements all of the above. Exposes the underlying `public readonly Spreadsheet $spreadsheet`, plus extras: `setStyle(string $range, array $style)` (keys: `bold`, `italic`, `size`, `underline`, `color`, `background_color`, `borders`, `borders_outline`, `align`, `valign`, `wrap`), `insertImage(string $filename, string $cell, array $options = [])`, `setCellFormat(string $range, string $format)`, `setWidth`, `setHeight`, `autoWidth(string $column)`, `findCell($value, bool $strict = false)`, `duplicateRows(...)`. Constants: `FORMAT_DATE`, `FORMAT_DATETIME`, `FORMAT_INT`, `FORMAT_DOUBLE`, `FORMAT_DOUBLE_10`, `FORMAT_PERCENTAGE`.
- `ZipDriver` - fast, low-memory writer; supports `Format::Xlsx` (grid) and `Format::Docx` (templates). Methods: `create()`, `load()`, `save()`, `replace()`, `setGrid()`, `setStyle(string $column, string $style)` where style is one of `header|integer|double|double_10|string|date|percentage|datetime`, `setWidth`, `setHeight`, `setSheetTitle`. Save format must equal load format.

## Usage examples

### Render an XLSX template (most common)

```php
use AnourValar\Office\SheetsService;
use AnourValar\Office\Format;

$data = [
    'vat' => 'No',
    'total' => ['price' => 2004.14, 'qty' => 3],
    'products' => [
        ['name' => 'Product #1', 'price' => 989,    'qty' => 1, 'date' => new \DateTime('2022-03-30')],
        ['name' => 'Product #2', 'price' => 1015.14,'qty' => 2, 'date' => new \DateTime('2022-03-31')],
    ],
];

(new SheetsService())
    ->generate(storage_path('templates/invoice.xlsx'), $data)
    ->saveAs(storage_path('app/invoice.xlsx'), Format::Xlsx);
```

### Stream a PDF download from a Laravel controller

```php
use AnourValar\Office\SheetsService;
use AnourValar\Office\Format;

$binary = (new SheetsService())
    ->generate(resource_path('templates/report.xlsx'), $data)
    ->save(Format::Pdf);

return response($binary, 200, [
    'Content-Type'        => Format::Pdf->contentType(),
    'Content-Disposition' => 'attachment; filename="report.pdf"',
]);
```

### Fill a DOCX template

```php
use AnourValar\Office\DocumentService;

(new DocumentService())
    ->generate(resource_path('templates/contract.docx'), [
        'client' => ['name' => 'Acme Inc.'],
        'total'  => 1234.56,
    ])
    ->saveAs(storage_path('app/contract.docx'));
```

### Simple grid export

```php
use AnourValar\Office\GridService;
use AnourValar\Office\Drivers\ZipDriver; // faster than the default PhpSpreadsheetDriver

(new GridService(new ZipDriver()))
    ->generate(['Name', 'Sales'], [
        ['William', 3000],
        ['James',   4000],
        ['Sveta',   5000],
    ])
    ->saveAs(storage_path('app/sales.xlsx'));
```

### Grid export with hooks and styling

```php
use AnourValar\Office\GridService;
use AnourValar\Office\Drivers\GridInterface;

$headers = [
    ['title' => 'Name', 'width' => 30],
    ['title' => 'Sales'],
];

$data = function () {
    foreach (Sale::query()->cursor() as $sale) {
        yield ['name' => $sale->name, 'sales' => $sale->amount];
    }
};

(new GridService())
    ->hookHeader(function (GridInterface $driver, mixed $header, $key, string $column) {
        if (isset($header['width'])) {
            $driver->setWidth($column, $header['width']);
        } else {
            $driver->autoWidth($column);
        }
        return $header['title'];
    })
    ->hookRow(fn (GridInterface $driver, mixed $row, $key) => [$row['name'], $row['sales']])
    ->hookAfter(function (GridInterface $driver, ?string $headersRange, ?string $dataRange, ?string $totalRange) {
        $driver->setSheetTitle('Sales');
        if ($headersRange) {
            $driver->setStyle($headersRange, ['bold' => true, 'background_color' => 'EEEEEE']);
        }
        if ($totalRange) {
            $driver->setStyle($totalRange, ['borders' => true, 'align' => 'left']);
        }
    })
    ->generate($headers, $data)
    ->saveAs(storage_path('app/sales.xlsx'));
```

### Eloquent-backed export via `ExportService` + `ExportGridQueryInterface`

```php
use AnourValar\Office\Facades\ExportService;
use AnourValar\Office\Facades\ExportGridQueryInterface;
use AnourValar\Office\Drivers\GridInterface;
use AnourValar\Office\Format;

class UsersGrid implements ExportGridQueryInterface
{
    public function query(): \Illuminate\Database\Eloquent\Builder
    {
        return \App\Models\User::query()->orderBy('id');
    }

    public function sheetTitle(array $request): string { return 'Users'; }

    public function columns(array $request): array
    {
        return [
            ['title' => 'ID',         'width' => 8],
            ['title' => 'Email',      'width' => 32],
            ['title' => 'Registered', 'width' => 20, ExportService::DATETIME => true],
        ];
    }

    public function item($row, GridInterface $driver, int $rowNumber, array $request): array
    {
        return [$row->id, $row->email, $row->created_at];
    }

    public function fileName(string $ext, array $request): string
    {
        return "users-".date('Y-m-d').".$ext";
    }
}

// Controller action
public function export(ExportService $exportService, UsersGrid $grid)
{
    $format = Format::Xlsx;
    $request = request()->all();

    return response()->streamDownload(
        function () use ($exportService, $grid, $format, $request) {
            echo $exportService->grid(
                fn () => yield from $grid->query()->cursor(),
                $grid,
                $format,
                $request,
            );
        },
        $grid->fileName($format->fileExtension(), $request),
        [
            'Content-Type'                    => $format->contentType(),
            'Access-Control-Expose-Headers'   => 'Content-Disposition',
        ],
    );
}
```

### Build a template dynamically with `hookLoad` + `hookBefore`

```php
use AnourValar\Office\SheetsService;
use AnourValar\Office\Drivers\SheetsInterface;

(new SheetsService())
    ->hookLoad(fn (SheetsInterface $driver) => $driver->create()) // start empty
    ->hookBefore(function (SheetsInterface $driver, array &$data) {
        $row = 1;
        foreach (array_keys($data) as $group) {
            $driver
                ->setValue("A$row", "[{$group}.name]")
                ->mergeCells("A$row:B$row");
            $row++;
            $driver
                ->setValue("A$row", "[$group.products.name]")
                ->setValue("B$row", "[$group.products.stock]");
            $row++;
        }
    })
    ->generate('', $data)
    ->saveAs(storage_path('app/dynamic.xlsx'));
```

### Merge two generated workbooks

```php
use AnourValar\Office\SheetsService;
use AnourValar\Office\Mixer;

$a = (new SheetsService())->generate('template.xlsx', ['foo' => 'hello']);
$b = (new SheetsService())->generate('template.xlsx', ['foo' => 'world']);

(new Mixer())($a, $b)->saveAs(storage_path('app/combined.xlsx'));
```

## Conventions / gotchas

- Template markers use square brackets, e.g. `[customer.name]`, `[products.name]` (one-dimensional table), and two-dimensional structures with `titles`/`values` keys. See README in the package root for visual examples.
- `SheetsService` infers the template format from the file extension (defaults to `Xlsx`). `DocumentService` defaults to `Docx`.
- `Generated::saveAs($filename)` will infer format from the filename extension; pass `Format` explicitly to override.
- `Format::Pdf` requires `mpdf/mpdf`; `Format::Docx` and the fast `ZipDriver` require `maennchen/zipstream-php`; everything else needs `phpoffice/phpspreadsheet`.
- `ZipDriver` only supports saving in the same format it loaded - it cannot convert XLSX to PDF, etc. Use `PhpSpreadsheetDriver` (the default) for conversions.
- `ExportService::grid()` automatically picks `ZipDriver` for `Format::Xlsx` (performance) and `PhpSpreadsheetDriver` otherwise.
- In `hookValue` / `hookRow`, returning `null` removes the cell/row from output.
- This package is framework-agnostic - there is no service provider, no facade alias, and no Laravel config to publish. To use Laravel DI, type-hint the service class (e.g. `SheetsService`, `ExportService`) in the constructor of your class.
- When using `ExportGridQueryInterface`, stream with `$grid->query()->cursor()` (or chunked) inside the data generator to keep memory low on large exports.
- For raw PhpSpreadsheet access inside hooks, use `$driver->spreadsheet` on `PhpSpreadsheetDriver` (it is a `public readonly` property).
