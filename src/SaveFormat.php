<?php

namespace AnourValar\Office;

enum SaveFormat: string
{
    case Xlsx = 'saveXlsx';
    case Pdf = 'savePdf';
    case Html = 'saveHtml';
}
