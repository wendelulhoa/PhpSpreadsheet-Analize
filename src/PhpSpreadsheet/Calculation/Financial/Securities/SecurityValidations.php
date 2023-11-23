<?php

namespace Analize\PhpSpreadsheet\Calculation\Financial\Securities;

use Analize\PhpSpreadsheet\Calculation\Exception;
use Analize\PhpSpreadsheet\Calculation\Financial\FinancialValidations;
use Analize\PhpSpreadsheet\Calculation\Information\ExcelError;

class SecurityValidations extends FinancialValidations
{
    public static function validateIssueDate(mixed $issue): float
    {
        return self::validateDate($issue);
    }

    public static function validateSecurityPeriod(mixed $settlement, mixed $maturity): void
    {
        if ($settlement >= $maturity) {
            throw new Exception(ExcelError::NAN());
        }
    }

    public static function validateRedemption(mixed $redemption): float
    {
        $redemption = self::validateFloat($redemption);
        if ($redemption <= 0.0) {
            throw new Exception(ExcelError::NAN());
        }

        return $redemption;
    }
}
