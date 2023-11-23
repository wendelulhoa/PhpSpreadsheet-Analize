<?php

namespace Analize\PhpSpreadsheet\Writer;

use Analize\PhpSpreadsheet\Spreadsheet;
use Analize\PhpSpreadsheet\Writer\Exception as WriterException;
use Analize\PhpSpreadsheet\Writer\Ods\Content;
use Analize\PhpSpreadsheet\Writer\Ods\Meta;
use Analize\PhpSpreadsheet\Writer\Ods\MetaInf;
use Analize\PhpSpreadsheet\Writer\Ods\Mimetype;
use Analize\PhpSpreadsheet\Writer\Ods\Settings;
use Analize\PhpSpreadsheet\Writer\Ods\Styles;
use Analize\PhpSpreadsheet\Writer\Ods\Thumbnails;
use ZipStream\Exception\OverflowException;
use ZipStream\ZipStream;

class Ods extends BaseWriter
{
    /**
     * Private PhpSpreadsheet.
     *
     * @var Spreadsheet
     */
    private $spreadSheet;

    private Content $writerPartContent;

    private Meta $writerPartMeta;

    private MetaInf $writerPartMetaInf;

    private Mimetype $writerPartMimetype;

    private Settings $writerPartSettings;

    private Styles $writerPartStyles;

    private Thumbnails $writerPartThumbnails;

    /**
     * Create a new Ods.
     */
    public function __construct(Spreadsheet $spreadsheet)
    {
        $this->setSpreadsheet($spreadsheet);

        $this->writerPartContent = new Content($this);
        $this->writerPartMeta = new Meta($this);
        $this->writerPartMetaInf = new MetaInf($this);
        $this->writerPartMimetype = new Mimetype($this);
        $this->writerPartSettings = new Settings($this);
        $this->writerPartStyles = new Styles($this);
        $this->writerPartThumbnails = new Thumbnails($this);
    }

    public function getWriterPartContent(): Content
    {
        return $this->writerPartContent;
    }

    public function getWriterPartMeta(): Meta
    {
        return $this->writerPartMeta;
    }

    public function getWriterPartMetaInf(): MetaInf
    {
        return $this->writerPartMetaInf;
    }

    public function getWriterPartMimetype(): Mimetype
    {
        return $this->writerPartMimetype;
    }

    public function getWriterPartSettings(): Settings
    {
        return $this->writerPartSettings;
    }

    public function getWriterPartStyles(): Styles
    {
        return $this->writerPartStyles;
    }

    public function getWriterPartThumbnails(): Thumbnails
    {
        return $this->writerPartThumbnails;
    }

    /**
     * Save PhpSpreadsheet to file.
     *
     * @param resource|string $filename
     */
    public function save($filename, int $flags = 0): void
    {
        $this->processFlags($flags);

        // garbage collect
        $this->spreadSheet->garbageCollect();

        $this->openFileHandle($filename);

        $zip = $this->createZip();

        $zip->addFile('META-INF/manifest.xml', $this->getWriterPartMetaInf()->write());
        $zip->addFile('Thumbnails/thumbnail.png', $this->getWriterPartthumbnails()->write());
        // Settings always need to be written before Content; Styles after Content
        $zip->addFile('settings.xml', $this->getWriterPartsettings()->write());
        $zip->addFile('content.xml', $this->getWriterPartcontent()->write());
        $zip->addFile('meta.xml', $this->getWriterPartmeta()->write());
        $zip->addFile('mimetype', $this->getWriterPartmimetype()->write());
        $zip->addFile('styles.xml', $this->getWriterPartstyles()->write());

        // Close file
        try {
            $zip->finish();
        } catch (OverflowException) {
            throw new WriterException('Could not close resource.');
        }

        $this->maybeCloseFileHandle();
    }

    /**
     * Create zip object.
     */
    private function createZip(): ZipStream
    {
        // Try opening the ZIP file
        if (!is_resource($this->fileHandle)) {
            throw new WriterException('Could not open resource for writing.');
        }

        // Create new ZIP stream
        return ZipStream0::newZipStream($this->fileHandle);
    }

    /**
     * Get Spreadsheet object.
     *
     * @return Spreadsheet
     */
    public function getSpreadsheet()
    {
        return $this->spreadSheet;
    }

    /**
     * Set Spreadsheet object.
     *
     * @param Spreadsheet $spreadsheet PhpSpreadsheet object
     *
     * @return $this
     */
    public function setSpreadsheet(Spreadsheet $spreadsheet): static
    {
        $this->spreadSheet = $spreadsheet;

        return $this;
    }
}
