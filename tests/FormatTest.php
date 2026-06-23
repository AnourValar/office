<?php

namespace AnourValar\Office\Tests;

use AnourValar\Office\Format;

class FormatTest extends \PHPUnit\Framework\TestCase
{
    /**
     * @return void
     */
    public function test_file_extension()
    {
        foreach (Format::cases() as $format) {
            $this->assertSame($format->value, $format->fileExtension());
        }

        $this->assertSame('xlsx', Format::Xlsx->fileExtension());
        $this->assertSame('pdf', Format::Pdf->fileExtension());
        $this->assertSame('html', Format::Html->fileExtension());
        $this->assertSame('ods', Format::Ods->fileExtension());
        $this->assertSame('csv', Format::Csv->fileExtension());
        $this->assertSame('docx', Format::Docx->fileExtension());
    }

    /**
     * @return void
     */
    public function test_content_type()
    {
        $this->assertSame(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            Format::Xlsx->contentType()
        );
        $this->assertSame('application/pdf', Format::Pdf->contentType());
        $this->assertSame('text/html', Format::Html->contentType());
        $this->assertSame('application/vnd.oasis.opendocument.spreadsheet', Format::Ods->contentType());
        $this->assertSame('text/csv', Format::Csv->contentType());
        $this->assertSame(
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            Format::Docx->contentType()
        );
    }
}
