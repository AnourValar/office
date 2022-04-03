<?php

namespace AnourValar\Office;

enum Format
{
    case Xlsx;
    case Pdf;
    case Html;
    case Ods;

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
