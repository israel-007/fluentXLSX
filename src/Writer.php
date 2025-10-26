<?php

namespace Fluentxlsx;

use Shuchkin\SimpleXLSXGen;

class Writer
{
    protected $xlsx;
    protected $sheets = [];
    protected $currentSheetName = 'Sheet1';

    public function __construct()
    {
        $this->xlsx = SimpleXLSXGen::fromArray([]);
        $this->sheets[$this->currentSheetName] = [];
    }

    /**
     * Add a new sheet (or switch to an existing one)
     */
    public function sheet($name)
    {
        $this->currentSheetName = $name;

        if (!isset($this->sheets[$name])) {
            $this->sheets[$name] = [];
        }

        return $this;
    }

    /**
     * Add a full row of data
     */
    public function addRow(array $row)
    {
        $this->sheets[$this->currentSheetName][] = $row;
        return $this;
    }

    /**
     * Add multiple rows
     */
    public function addRows(array $rows)
    {
        foreach ($rows as $row) {
            $this->addRow($row);
        }
        return $this;
    }

    /**
     * Set single cell value by Excel reference (e.g. "A1", "B2")
     */
    public function cell($ref, $value)
    {
        if (!is_string($ref) || !preg_match('/^\s*([A-Z]+)([0-9]+)\s*$/i', $ref, $matches)) {
            throw new \InvalidArgumentException("Invalid cell reference: $ref");
        }

        $colLetters = strtoupper($matches[1]);
        $row = (int) $matches[2];
        $col = $this->columnNameToIndex($colLetters);

        return $this->cellEx($row, $col, $value);
    }

    /**
     * Set single cell by row and col
     */
    public function cellEx($row, $col, $value)
    {
        if (is_string($col)) {
            $col = $this->columnNameToIndex($col);
        }

        $this->setCell($row, $col, $value);
        return $this;
    }

    /**
     * Internal: set a cell ensuring rows/columns are normalized (fills gaps with nulls)
     */
    protected function setCell($row, $col, $value)
    {
        $rIndex = max(0, (int) $row - 1);
        $cIndex = max(0, (int) $col - 1);

        // Ensure sheet exists
        if (!isset($this->sheets[$this->currentSheetName]) || !is_array($this->sheets[$this->currentSheetName])) {
            $this->sheets[$this->currentSheetName] = [];
        }

        // Ensure rows up to rIndex exist
        for ($i = 0; $i <= $rIndex; $i++) {
            if (!isset($this->sheets[$this->currentSheetName][$i]) || !is_array($this->sheets[$this->currentSheetName][$i])) {
                $this->sheets[$this->currentSheetName][$i] = [];
            }
        }

        // Ensure columns up to cIndex exist for this row
        for ($j = 0; $j <= $cIndex; $j++) {
            if (!array_key_exists($j, $this->sheets[$this->currentSheetName][$rIndex])) {
                $this->sheets[$this->currentSheetName][$rIndex][$j] = null;
            }
        }

        // Set the value
        $this->sheets[$this->currentSheetName][$rIndex][$cIndex] = $value;
    }

    /**
     * Save to file
     */
    public function save($filename)
    {
        $this->build();
        $this->xlsx->saveAs($filename);
        return $this;
    }

    /**
     * Download directly
     */
    public function download($filename = 'export.xlsx')
    {
        $this->build();
        $this->xlsx->downloadAs($filename);
        return $this;
    }

    /**
     * Build XLSX from collected sheets (normalizes rows/columns before passing to SimpleXLSXGen)
     */
    protected function build()
    {
        $normalizedSheets = [];

        foreach ($this->sheets as $name => $rows) {
            if (empty($rows)) {
                $normalizedSheets[$name] = [];
                continue;
            }

            $rowKeys = array_keys($rows);
            $maxRowIndex = max($rowKeys);
            $normalized = [];

            for ($i = 0; $i <= $maxRowIndex; $i++) {
                $row = (isset($rows[$i]) && is_array($rows[$i])) ? $rows[$i] : [];

                // Make sure columns are numeric and continuous from 0..N
                ksort($row);
                $colKeys = array_keys($row);
                $maxColIndex = empty($colKeys) ? -1 : max($colKeys);

                $newRow = [];
                for ($j = 0; $j <= $maxColIndex; $j++) {
                    $newRow[$j] = array_key_exists($j, $row) ? $row[$j] : null;
                }

                // Ensure row is a zero-indexed sequential array
                $normalized[] = array_values($newRow);
            }

            $normalizedSheets[$name] = $normalized;
        }

        // Build SimpleXLSXGen object
        $names = array_keys($normalizedSheets);
        if (empty($names)) {
            $this->xlsx = SimpleXLSXGen::fromArray([]);
            return;
        }

        $first = array_shift($names);
        $this->xlsx = SimpleXLSXGen::fromArray($normalizedSheets[$first], $first);

        foreach ($names as $name) {
            $this->xlsx->addSheet($normalizedSheets[$name], $name);
        }
    }

    public function __call($method, $args)
    {
        if (method_exists($this->xlsx, $method)) {
            return call_user_func_array([$this->xlsx, $method], $args);
        }

        throw new \BadMethodCallException("Method $method does not exist in Writer or SimpleXLSXGen");
    }

    protected function columnNameToIndex($letters)
    {
        $letters = strtoupper($letters);
        $len = strlen($letters);
        $index = 0;

        for ($i = 0; $i < $len; $i++) {
            $index = $index * 26 + (ord($letters[$i]) - ord('A') + 1);
        }

        return $index;
    }
}
