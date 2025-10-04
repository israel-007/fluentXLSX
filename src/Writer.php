<?php

namespace Fluentxlsx;

use Shuchkin\SimpleXLSXGen;

class Writer
{
    protected array $sheets = [];

    /**
     * Add a new sheet with data.
     *
     * @param string $name
     * @param array $data
     * @return $this
     */
    public function sheet(string $name, array $data): self
    {
        $this->sheets[$name] = $data;
        return $this;
    }

    /**
     * Save the Excel file to disk.
     *
     * @param string $filename
     * @return void
     */
    public function save(string $filename): void
    {
        $xlsx = $this->build();
        $xlsx->saveAs($filename);
    }

    /**
     * Stream the Excel file to browser (download).
     *
     * @param string $filename
     * @return void
     */
    public function download(string $filename): void
    {
        $xlsx = $this->build();
        $xlsx->downloadAs($filename);
    }

    /**
     * Build the SimpleXLSXGen instance with all sheets.
     *
     * @return SimpleXLSXGen
     */
    protected function build(): SimpleXLSXGen
    {
        $xlsx = SimpleXLSXGen::fromArray([]);

        foreach ($this->sheets as $name => $data) {
            $xlsx->addSheet($data, $name);
        }

        return $xlsx;
    }
}
