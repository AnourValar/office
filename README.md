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

**template.xlsx:**

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
    ->render(
        'template.xlsx', // path to template
        $data, // input data
        \AnourValar\Office\SaveFormat::Xlsx // save format
    )
    ->save('generated_document.xlsx');

// Available save formats:
// \AnourValar\Office\SaveFormat::Xlsx
// \AnourValar\Office\SaveFormat::Pdf
// \AnourValar\Office\SaveFormat::Html
// \AnourValar\Office\SaveFormat::Ods
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-11.png)


**The same template with empty data**

![Demo](https://anour.ru/resources/office-v1-12.png)


### Two-dimensional table

**template.xlsx:**

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
    ->render('template.xlsx' $data)
    ->save('generated_document.xlsx');
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office-v1-21.png)

