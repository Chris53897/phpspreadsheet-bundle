<?php

namespace Yectep\PhpSpreadsheetBundle;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\BaseWriter;
use PhpOffice\PhpSpreadsheet\Writer\Exception;
use PhpOffice\PhpSpreadsheet\Writer\IWriter;
use Symfony\Component\HttpFoundation\StreamedResponse;

/**
 * Factory class for PhpSpreadsheet objects.
 *
 * @package Yectep\PhpSpreadsheetBundle
 */
class Factory {

    /**
     * Returns a new instance of the PhpSpreadsheet class.
     *
     * @param string|null $filename     If set, uses the IOFactory to return the spreadsheet located at $filename
     *                                  using automatic type resolution per \PhpOffice\PhpSpreadsheet\IOFactory.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function createSpreadsheet(?string $filename = null): Spreadsheet
    {
        return (is_null($filename) ? new Spreadsheet() : IOFactory::load($filename));
    }

    /**
     * Returns the PhpSpreadsheet IWriter instance to save a file.
     *
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function createWriter(Spreadsheet $spreadsheet, string $type): IWriter
    {
        return IOFactory::createWriter($spreadsheet, $type);
    }

    /**
     * @param string $type   Reader class to create.
     *
     * @return mixed            Returns a IReader of the given type if found.
     * @throws \InvalidArgumentException
     */
    public function createReader(string $type): mixed
    {
        $readerClass = '\\PhpOffice\\PhpSpreadsheet\\Reader\\' . $type;
        if (!class_exists($readerClass)) {
            throw new \InvalidArgumentException('The reader [' . $type . '] does not exist or is not supported by PhpSpreadsheet.');
        }

        return new $readerClass();
    }


    /**
     * Return a StreamedResponse containing the file
     * @throws Exception
     */
    public function createStreamedResponse(Spreadsheet $spreadsheet, string $type, int $status = 200, array $headers = [], array $writerOptions = []): StreamedResponse
    {
        $writer = IOFactory::createWriter($spreadsheet, $type);

        if (!empty($writerOptions)) {
            $this->applyOptionsToWriter($writer, $writerOptions);
        }

        return new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            },
            $status,
            $headers
        );
    }

    /**
     * @param array $options
     */
    private function applyOptionsToWriter(BaseWriter $writer, array $options = []): void
    {
        foreach ($options as $method => $arguments) {
            if (method_exists($writer, $method)) {
                if (!is_array($arguments)) {
                    $arguments = array($arguments);
                }
                call_user_func_array(array($writer, $method), $arguments);
            }
        }
    }
}
