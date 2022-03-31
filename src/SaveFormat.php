<?php

namespace AnourValar\Office;

enum SaveFormat
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
            SaveFormat::Xlsx => 'xlsx',
            SaveFormat::Pdf => 'pdf',
            SaveFormat::Html => 'html',
            SaveFormat::Ods => 'ods',
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
            SaveFormat::Xlsx => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            SaveFormat::Pdf => 'application/pdf',
            SaveFormat::Html => 'text/html',
            SaveFormat::Ods => 'application/vnd.oasis.opendocument.spreadsheet',
        };
    }

    /**
     * @see \AnourValar\Office\Drivers\TemplateInterface
     *
     * @return string
     */
    public function driverSaveMethod(): string
    {
        return match($this)
        {
            SaveFormat::Xlsx => 'saveXlsx',
            SaveFormat::Pdf => 'savePdf',
            SaveFormat::Html => 'saveHtml',
            SaveFormat::Ods => 'saveOds',
        };
    }
}
