# Office: Documents & Reports

## Installation

### Minimal

```bash
composer require anourvalar/office
```

### To work with default driver - Phpspreadsheet is required:

```bash
composer require phpoffice/phpspreadsheet "^1.18"
```


### To save document as PDF - Mpdf is required:

```bash
composer require mpdf/mpdf: "^8.0"
```


## Usage

### Generate a document from an XLSX (Excel) template

**template.xlsx:**

![Demo](https://anour.ru/resources/office_template.xlsx.png)



```php
$data = [
    'vat' => 'No',
    'total' => [
        'price' => 2004.14,
        'qty' => 3,
    ],

    'products' => [ // list
        [
            'name' => 'Product #1',
            'price' => 989,
            'qty' => 1,
            'date' => new DateTime('2022-03-30'),
        ],
        [
            'name' => 'Product #2',
            'price' => 1015.14,
            'qty' => 2,
            'date' => new DateTime('2022-03-31'),
        ],
    ],

    'managers' => [ // matrix (multidimensional)
        'names' => [[ 'William', 'James', 'Svetlana' ]],
        'stats' => [
            [
                'month' => 'January',
                'amount' => [700, 800, 900],
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
    ->render('template.xlsx', $data)
    ->save('generated_document.xlsx');

// Save as PDF
(new \AnourValar\Office\TemplateService())
    ->render('template.xlsx', $data, \AnourValar\Office\SaveFormat::Pdf)
    ->save('generated_document.pdf');

// Save as HTML
(new \AnourValar\Office\TemplateService())
    ->render('template.xlsx', $data, \AnourValar\Office\SaveFormat::Html)
    ->save('generated_document.html');
```

**generated_document.xlsx:**

![Demo](https://anour.ru/resources/office_generated_document.xlsx.png)
