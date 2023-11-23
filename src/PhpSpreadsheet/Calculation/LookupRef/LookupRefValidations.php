<?php

namespace Analize\PhpSpreadsheet\Calculation\LookupRef;

use Analize\PhpSpreadsheet\Calculation\Exception;
use Analize\PhpSpreadsheet\Calculation\Information\ErrorValue;
use Analize\PhpSpreadsheet\Calculation\Information\ExcelError;

class LookupRefValidations
{
    public static function validateInt(mixed $value): int
    {
        if (!is_numeric($value)) {
            if (ErrorValue::isError($value)) {
                throw new Exception($value);
            }

            throw new Exception(ExcelError::VALUE());
        }

        return (int) floor((float) $value);
    }

    public static function validatePositiveInt(mixed $value, bool $allowZero = true): int
    {
        $value = self::validateInt($value);

        if (($allowZero === false && $value <= 0) || $value < 0) {
            throw new Exception(ExcelError::VALUE());
        }

        return $value;
    }
}
