<?php

namespace Analize\PhpSpreadsheet\Writer\Xls;

use Analize\PhpSpreadsheet\Shared\StringHelper;

class Font
{
    /**
     * Color index.
     */
    private int $colorIndex;

    /**
     * Font.
     */
    private \Analize\PhpSpreadsheet\Style\Font $font;

    /**
     * Constructor.
     */
    public function __construct(\Analize\PhpSpreadsheet\Style\Font $font)
    {
        $this->colorIndex = 0x7FFF;
        $this->font = $font;
    }

    /**
     * Set the color index.
     *
     * @param int $colorIndex
     */
    public function setColorIndex($colorIndex): void
    {
        $this->colorIndex = $colorIndex;
    }

    /** @var int */
    private static $notImplemented = 0;

    /**
     * Get font record data.
     */
    public function writeFont(): string
    {
        $font_outline = self::$notImplemented;
        $font_shadow = self::$notImplemented;

        $icv = $this->colorIndex; // Index to color palette
        if ($this->font->getSuperscript()) {
            $sss = 1;
        } elseif ($this->font->getSubscript()) {
            $sss = 2;
        } else {
            $sss = 0;
        }
        $bFamily = 0; // Font family
        $bCharSet = \Analize\PhpSpreadsheet\Shared\Font::getCharsetFromFontName((string) $this->font->getName()); // Character set

        $record = 0x31; // Record identifier
        $reserved = 0x00; // Reserved
        $grbit = 0x00; // Font attributes
        if ($this->font->getItalic()) {
            $grbit |= 0x02;
        }
        if ($this->font->getStrikethrough()) {
            $grbit |= 0x08;
        }
        if ($font_outline) {
            $grbit |= 0x10;
        }
        if ($font_shadow) {
            $grbit |= 0x20;
        }

        $data = pack(
            'vvvvvCCCC',
            // Fontsize (in twips)
            $this->font->getSize() * 20,
            $grbit,
            // Colour
            $icv,
            // Font weight
            self::mapBold($this->font->getBold()),
            // Superscript/Subscript
            $sss,
            self::mapUnderline((string) $this->font->getUnderline()),
            $bFamily,
            $bCharSet,
            $reserved
        );
        $data .= StringHelper::UTF8toBIFF8UnicodeShort((string) $this->font->getName());

        $length = strlen($data);
        $header = pack('vv', $record, $length);

        return $header . $data;
    }

    /**
     * Map to BIFF5-BIFF8 codes for bold.
     */
    private static function mapBold(?bool $bold): int
    {
        if ($bold === true) {
            return 0x2BC; //  700 = Bold font weight
        }

        return 0x190; //  400 = Normal font weight
    }

    /**
     * Map of BIFF2-BIFF8 codes for underline styles.
     *
     * @var int[]
     */
    private static $mapUnderline = [
        \Analize\PhpSpreadsheet\Style\Font::UNDERLINE_NONE => 0x00,
        \Analize\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLE => 0x01,
        \Analize\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLE => 0x02,
        \Analize\PhpSpreadsheet\Style\Font::UNDERLINE_SINGLEACCOUNTING => 0x21,
        \Analize\PhpSpreadsheet\Style\Font::UNDERLINE_DOUBLEACCOUNTING => 0x22,
    ];

    /**
     * Map underline.
     *
     * @param string $underline
     *
     * @return int
     */
    private static function mapUnderline($underline)
    {
        if (isset(self::$mapUnderline[$underline])) {
            return self::$mapUnderline[$underline];
        }

        return 0x00;
    }
}
