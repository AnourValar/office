<?php

namespace AnourValar\Office;

enum Format: string
{
    case Xlsx = 'xlsx'; // sheets | grid => reader + write
    case Pdf = 'pdf'; // sheets | grid => writer
    case Html = 'html'; // sheets | grid => reader + write
    case Ods = 'ods'; // sheets | grid => reader + write

    /**
     * @return string
     */
    public function fileExtension(): string
    {
        return match($this)
        {
            Format::Xlsx => 'xlsx',
            Format::Pdf => 'pdf',
            Format::Html => 'html',
            Format::Ods => 'ods',
        };
    }

    /**
     * MIME
     *
     * @return string
     */
    public function contentType(): string
    {
        return match($this)
        {
            Format::Xlsx => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            Format::Pdf => 'application/pdf',
            Format::Html => 'text/html',
            Format::Ods => 'application/vnd.oasis.opendocument.spreadsheet',
        };
    }
}
