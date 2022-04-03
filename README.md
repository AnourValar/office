# Office: Documents & Reports

## Installation

### Minimal

```bash
composer require anourvalar/office
```

### To work with default driver - Phpspreadsheet is required:

```bash
composer require phpoffice/phpspreadsheet "^1.22"
```


### To save documents as PDF - Mpdf is required:

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
(new \AnourValar\Office\TemplateService())
    ->generate(
        'template1.xlsx', // path to template
        $data // input data
    )
    ->saveToFile(
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
(new \AnourValar\Office\TemplateService())
    ->generate('template2.xlsx', $data)
    ->saveToFile('generated_document.xlsx'); // xlsx as save format is set by default
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-21.png)

### Additional Features

**template3.xlsx:**

![Demo](https://anour.ru/resources/office-v1-30.png)

```php
$data = [
    'foo' => 'Hello',
    
    'bar' => function (\AnourValar\Office\Drivers\TemplateInterface $driver, $cell) {
        $driver->insertImage('logo.png', $cell, ['width' => 100, 'offset_y' => -45]);
        return 'Logo!'; // replace marker "[bar]" with return value
    }
];

(new \AnourValar\Office\TemplateService())
    ->hookValue(function (TemplateInterface $driver, $cell, $value) {
        // Hook will be called for every cell's value which is changing
    
        $value .= ' world';
        return $value;
    })
    ->generate(
        'template3.ods',
        $data,
        \AnourValar\Office\Format::Ods // template's format
    )
    ->saveToFile('generated_document.xlsx');

// Available hooks:
// hookLoad: Closure(TemplateInterface $driver, string $templateFile, Format $templateFormat)
// hookBefore: Closure(TemplateInterface $templateDriver, array &$data)
// hookValue: Closure(TemplateInterface $templateDriver, string $cell, mixed $value)
// hookAfter: Closure(TemplateInterface $templateDriver)
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-31.png)

### Merge (union)

```php
$dataA = ['foo' => 'hello'];
$dataB = ['foo' => 'world'];

$documentA = (new \AnourValar\Office\TemplateService())->generate('template.xlsx', $dataA);
$documentB = (new \AnourValar\Office\TemplateService())->generate('template.xlsx', $dataB);

$mixer = new \AnourValar\Office\MixerService();
$mixer($documentA, $documentB)->saveToFile('generated_document.xlsx');
```
