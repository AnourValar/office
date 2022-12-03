<?php

namespace AnourValar\Office\Traits;

trait XFormat
{
    /**
     * @param \DateTimeInterface $date
     * @return float
     */
    protected function excelDate(\DateTimeInterface $date): float
    {
        $year = (int) $date->format('Y');
        $month = (int) $date->format('m');
        $day = (int) $date->format('d');
        $hours = (int) $date->format('H');
        $minutes = (int) $date->format('i');
        $seconds = (int) $date->format('s');

        $leapYear = true;
        if ($year == 1900 && $month <= 2) {
            $leapYear = false;
        }
        $baseDate = 2415020;

        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            --$year;
        }

        $century = (int) substr($year, 0, 2);
        $decade = (int) substr($year, 2, 2);
        $excelDate = floor((146097 * $century) / 4) + floor((1461 * $decade) / 4) + floor((153 * $month + 2) / 5);
        $excelDate += $day + 1721119 - $baseDate + $leapYear;
        $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;

        return (float) $excelDate + $excelTime;
    }

    /**
     * @param string|null $value
     * @return string
     */
    protected function escape(?string $value): string
    {
        return htmlspecialchars((string) $value, ENT_QUOTES, 'UTF-8', true);
    }
}
