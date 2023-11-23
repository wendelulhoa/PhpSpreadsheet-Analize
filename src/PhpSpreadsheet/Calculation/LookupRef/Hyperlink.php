<?php

namespace Analize\PhpSpreadsheet\Calculation\LookupRef;

use Analize\PhpSpreadsheet\Calculation\Functions;
use Analize\PhpSpreadsheet\Calculation\Information\ExcelError;
use Analize\PhpSpreadsheet\Cell\Cell;

class Hyperlink
{
    /**
     * HYPERLINK.
     *
     * Excel Function:
     *        =HYPERLINK(linkURL, [displayName])
     *
     * @param mixed $linkURL Expect string. Value to check, is also the value returned when no error
     * @param mixed $displayName Expect string. Value to return when testValue is an error condition
     * @param Cell $cell The cell to set the hyperlink in
     *
     * @return mixed The value of $displayName (or $linkURL if $displayName was blank)
     */
    public static function set(mixed $linkURL = '', mixed $displayName = null, ?Cell $cell = null)
    {
        $linkURL = ($linkURL === null) ? '' : Functions::flattenSingleValue($linkURL);
        $displayName = ($displayName === null) ? '' : Functions::flattenSingleValue($displayName);

        if ((!is_object($cell)) || (trim($linkURL) == '')) {
            return ExcelError::REF();
        }

        if ((is_object($displayName)) || trim($displayName) == '') {
            $displayName = $linkURL;
        }

        $cell->getHyperlink()->setUrl($linkURL);
        $cell->getHyperlink()->setTooltip($displayName);

        return $displayName;
    }
}
