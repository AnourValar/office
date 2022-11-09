# Office: Documents | Reports | Grids

## Installation

### Minimal

```bash
composer require anourvalar/office
```

### Phpspreadsheet is required to work with Excel (xlsx).

```bash
composer require phpoffice/phpspreadsheet "^1.25"
```

### Zipstream-php is required to work with Word (docx).

```bash
composer require maennchen/zipstream-php "^2.2"
```

### Mpdf is required to work with PDF.

```bash
composer require mpdf/mpdf: "^8.0"
```


## Generate a document from an XLSX (Excel) template

### One-dimensional table

**template1.xlsx:**

![Demo](https://anour.ru/resources/office-v1-10.png)

```php
$data = [
    // scalar
    'vat' => 'No',
    'total' => [
        'price' => 2004.14,
        'qty' => 3,
    ],

    // one-dimensional table
    'products' => [
        [
            'name' => 'Product #1',
            'price' => 989,
            'qty' => 1,
            'date' => new \DateTime('2022-03-30'),
        ],
        [
            'name' => 'Product #2',
            'price' => 1015.14,
            'qty' => 2,
            'date' => new \DateTime('2022-03-31'),
        ],
    ],
];

// Save as XLSX (Excel)
(new \AnourValar\Office\SheetsService())
    ->generate(
        'template1.xlsx', // path to template
        $data // input data
    )
    ->saveAs(
        'generated_document.xlsx', // path to save
        \AnourValar\Office\Format::Xlsx // save format
    );

// Available formats:
// \AnourValar\Office\Format::Xlsx
// \AnourValar\Office\Format::Pdf
// \AnourValar\Office\Format::Html
// \AnourValar\Office\Format::Ods
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-11.png)


**The same template with empty data**

![Demo](https://anour.ru/resources/office-v1-12.png)


### Two-dimensional table

**template2.xlsx:**

![Demo](https://anour.ru/resources/office-v1-20.png)

```php
$data = [
    'best_manager' => 'Sveta',

    // two-dimensional table
    'managers' => [
        'titles' => [[ 'William', 'James', 'Sveta' ]],

        'values' => [
            [ // additional row
                'month' => 'January',
                'amount' => [700, 800, 900], // additional columns
            ],
            [
                'month' => 'February',
                'amount' => [7000, 8000, 9000],
            ],
            [
                'month' => 'March',
                'amount' => [70000, 80000, 90000],
            ],
        ],
    ],
];

// Save as XLSX (Excel)
(new \AnourValar\Office\SheetsService())
    ->generate('template2.xlsx', $data)
    ->saveAs('generated_document.xlsx'); // second argument (format) is optional
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-21.png)

### Additional Features

**template3.xlsx:**

![Demo](https://anour.ru/resources/office-v1-30.png)

```php
$data = [
    'foo' => 'Hello',

    'bar' => function (\AnourValar\Office\Drivers\SheetsInterface $driver, $cell) {
        $driver->insertImage('logo.png', $cell, ['width' => 100, 'offset_y' => -45]);
        return 'Logo!'; // replace marker "[bar]" with return value
    }
];

(new \AnourValar\Office\SheetsService())
    ->hookValue(function (SheetsInterface $driver, $cell, $value, int $sheetIndex) {
        // Hook will be called for every cell which is changing

        $value .= ' world';
        return $value;
    })
    ->generate(
        'template3.ods', // ods template
        $data,
        true // cells auto format instead of template setup
    )
    ->saveAs('generated_document.xlsx');

// Available hooks:
// hookLoad: Closure(SheetsInterface $driver, string $templateFile, Format $templateFormat)
// hookBefore: Closure(SheetsInterface $driver, array &$data)
// hookValue: Closure(SheetsInterface $driver, string $cell, mixed $value, int $sheetIndex)
// hookAfter: Closure(SheetsInterface $driver)
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-31.png)

### Dynamic templates

```php
$data = [
    'group1' => [
        'name' => 'Group 1',
        'products' => [
            ['name' => 'Product 1', 'stock' => 101],
            ['name' => 'Product 2', 'stock' => 102],
        ],
    ],
    'group2' => [
        'name' => 'Group 2',
        'products' => [
            ['name' => 'Product 3', 'stock' => 103],
            ['name' => 'Product 4', 'stock' => 104],
        ],
    ],
];

(new \AnourValar\Office\SheetsService())
    ->hookLoad(function ($driver, string $templateFile, $templateFormat) {
        // create empty document instead of using existing
        return $driver->create();
    })
    ->hookBefore(function ($driver, array &$data) {
        // place markers on-fly
        $row = 1;
        foreach (array_keys($data) as $group) {
            // group's title
            $driver
                ->setValue("A$row", "[{$group}.name]")
                ->mergeCells("A$row:B$row")
                ->setStyle("A$row", ['align' => 'center', 'bold' => true]);
            $row++;

            // group's products
            $driver
                ->setValue("A$row", "[$group.products.name]")
                ->setValue("B$row", "[$group.products.stock]");
            $row++;
        }
    })
    ->generate('', $data)
    ->saveAs('generated_document.xlsx');
```

**Dynamic template overview**

![Demo](https://anour.ru/resources/office-v1-61.png)

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-62.png)

### Merge (union) few documents to a single file

```php
$dataA = ['foo' => 'hello'];
$dataB = ['foo' => 'world'];

$documentA = (new \AnourValar\Office\SheetsService())->generate('template.xlsx', $dataA);
$documentB = (new \AnourValar\Office\SheetsService())->generate('template.xlsx', $dataB);

$mixer = new \AnourValar\Office\Mixer();
$mixer($documentA, $documentB)->saveAs('generated_document.xlsx');
```


## Generate a document from an DOCX (Word) template

```php
(new \AnourValar\Office\DocumentService)
    ->generate('template.docx', ['foo' => 'bar'])
    ->saveAs('generated_document.docx');
```

**template.docx:**

![Demo](https://anour.ru/resources/office-v1-70.png)

**generated_document.docx:**

![Demo](https://anour.ru/resources/office-v1-71.png)


## Export table (Grid)

### Simple usage

```php
$data = [
    ['William', 3000],
    ['James', 4000],
    ['Sveta', 5000],
];

// Save as XLSX (Excel)
(new \AnourValar\Office\GridService())
    ->generate(
        ['Name', 'Sales'], // headers
        $data // data
    )
    ->saveAs('generated_grid.xlsx');
```

**generated_grid.xlsx:**

![Demo](https://anour.ru/resources/office-v1-41.png)

### Advanced usage

```php
$headers = [
    ['title' => 'Name', 'width' => 30],
    ['title' => 'Sales'],
];

$data = function () {
    yield ['name' => 'William', 'sales' => 3000];
    yield ['name' => 'James', 'sales' => 4000];
    yield ['name' => 'Sveta', 'sales' => 5000];
};

// Save as XLSX (Excel)
(new \AnourValar\Office\GridService())
    ->hookHeader(function (GridInterface $driver, mixed $header, $key, $column) {
        if (isset($header['width'])) {
            $driver->setWidth($column, $header['width']); // column with fixed width
        } else {
            $driver->autoWidth($column); // column with auto width
        }

        return $header['title'];
    })
    ->hookRow(function (GridInterface $driver, mixed $row, $key) {
        return [
            $row['name'],
            $row['sales'],
        ];
    })
    ->hookAfter(function (
        GridInterface $driver,
        string $headersRange,
        string $dataRange,
        string $totalRange,
        array $columns
    ) {
        $driver->setSheetTitle('Foo');

        $driver->setStyle(
            $headersRange, // A1:B1
            ['bold' => true, 'background_color' => 'EEEEEE']
        );

        $driver->setStyle(
            $totalRange, // A1:B4
            ['borders' => true, 'align' => 'left']
        );
    })
    ->generate($headers, $data)
    ->saveAs('generated_grid.xlsx');
```

**generated_grid.xlsx:**

![Demo](https://anour.ru/resources/office-v1-51.png)
