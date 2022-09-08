<?php

namespace AnourValar\Office;

enum Format: string
{
    case Xlsx = 'xlsx'; // sheets | grid => reader + writer
    case Pdf = 'pdf'; // sheets | grid => writer
    case Html = 'html'; // sheets | grid => reader + writer
    case Ods = 'ods'; // sheets | grid => reader + writer
    case Docx = 'docx'; // document => reader + writer

    /**
     * @return string
     */
    public function fileExtension(): string
    {
        return match($this) {
            Format::Xlsx => 'xlsx',
            Format::Pdf => 'pdf',
            Format::Html => 'html',
            Format::Ods => 'ods',
            Format::Docx => 'docx',
        };
    }

    /**
     * MIME
     *
     * @return string
     */
    public function contentType(): string
    {
        return match($this) {
            Format::Xlsx => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            Format::Pdf => 'application/pdf',
            Format::Html => 'text/html',
            Format::Ods => 'application/vnd.oasis.opendocument.spreadsheet',
            Format::Docx => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        };
    }
}
