<?php

namespace Fluentxlsx;

use Shuchkin\SimpleXLSX;

class Reader
{
    protected $xlsx;
    protected $currentSheetIndex = 0;
    protected $sheetExplicitlySet = false;

    public function fromFile($file)
    {
        if (!file_exists($file)) {
            throw new \Exception("File not found: $file");
        }

        $this->xlsx = SimpleXLSX::parse($file);

        if (!$this->xlsx) {
            throw new \Exception("Failed to parse Excel file: " . SimpleXLSX::parseError());
        }

        return $this;
    }

    public function fromData($data)
    {
        $this->xlsx = SimpleXLSX::parseData($data);

        if (!$this->xlsx) {
            throw new \Exception("Failed to parse Excel data: " . SimpleXLSX::parseError());
        }

        return $this;
    }

    /**
     * Select sheet by index, name, or "ACTIVE"
     */
    public function sheet($sheet)
    {
        $this->sheetExplicitlySet = true;

        if (is_int($sheet)) {
            if ($sheet < 0 || $sheet >= count($this->xlsx->sheetNames())) {
                throw new \Exception("Invalid sheet index: $sheet");
            }
            $this->currentSheetIndex = $sheet;
        } elseif (is_string($sheet)) {
            if (strtoupper($sheet) === "ACTIVE") {
                $this->currentSheetIndex = $this->xlsx->activeSheet;
            } else {
                $names = $this->xlsx->sheetNames();
                $index = array_search($sheet, $names, true);

                if ($index === false) {
                    throw new \Exception("Sheet not found with name: $sheet");
                }
                $this->currentSheetIndex = $index;
            }
        } else {
            throw new \Exception("Sheet identifier must be an integer, a string name, or 'ACTIVE'");
        }

        return $this;
    }

    /**
     * Get rows with optional limits
     *
     * @param int|array|null $limit
     *  - int: limit number of rows
     *  - array [start, end]: inclusive range of rows
     *  - null: all rows
     * @return array
     */
    public function rows($limit = null)
    {
        $this->ensureDefaultSheet();
        $allRows = $this->xlsx->rows($this->currentSheetIndex);

        if ($limit === null) {
            return $allRows;
        }

        // Case: rows(5) → first 5 rows
        if (is_int($limit)) {
            return array_slice($allRows, 0, $limit);
        }

        // Case: rows([2, 10]) → rows 2 to 10 (1-based index)
        if (is_array($limit) && count($limit) === 2) {
            $start = max(1, $limit[0]);
            $end = $limit[1];

            return array_slice($allRows, $start - 1, $end - $start + 1);
        }

        throw new \Exception("Invalid argument for rows(). Use int or [start, end]");
    }

    /**
     * Get explicit row numbers
     *
     * @param array $rows Array of row numbers (1-based)
     * @return array
     */
    public function rowsEx(array $rows)
    {
        $this->ensureDefaultSheet();
        $allRows = $this->xlsx->rows($this->currentSheetIndex);

        $result = [];
        foreach ($rows as $rowNum) {
            $index = $rowNum - 1;
            if (isset($allRows[$index])) {
                $result[] = $allRows[$index];
            }
        }

        return $result;
    }

    /**
     * Get headers of current sheet
     */
    public function headers()
    {
        $this->ensureDefaultSheet();
        return $this->xlsx->rows($this->currentSheetIndex)[0] ?? [];
    }

    /**
     * Convert sheet into associative array (headers as keys)
     */
    public function toAssoc()
    {
        $this->ensureDefaultSheet();

        $rows = $this->xlsx->rows($this->currentSheetIndex);
        if (count($rows) < 2) {
            return [];
        }

        $headers = array_shift($rows);
        $assoc = [];

        foreach ($rows as $row) {
            $assoc[] = array_combine($headers, $row);
        }

        return $assoc;
    }

    /**
     * Get all sheet names
     */
    public function sheets()
    {
        return $this->xlsx->sheetNames();
    }

    /**
     * Ensure we have a default sheet if none explicitly set
     */
    protected function ensureDefaultSheet()
    {
        if (!$this->sheetExplicitlySet) {
            $this->currentSheetIndex = 0;
            $this->sheetExplicitlySet = true;
        }
    }

    /**
     * Get cell by Excel-style reference (e.g. "A1", "AB6")
     *
     * @param string $ref
     * @return string|null
     * @throws \Exception
     */
    public function cell($ref)
    {
        $this->ensureDefaultSheet();

        if (!preg_match('/^([A-Z]+)([0-9]+)$/i', $ref, $matches)) {
            throw new \Exception("Invalid cell reference: $ref");
        }

        $colLetters = strtoupper($matches[1]);
        $row = (int) $matches[2];

        $col = $this->columnNameToIndex($colLetters);

        return $this->cellEx($row, $col);
    }

    /**
     * Get cell by row and column
     *
     * @param int $row 1-based row number
     * @param int|string $col Column index (1-based) or Excel letters ("A", "AB")
     * @return string|null
     * @throws \Exception
     */
    public function cellEx($row, $col)
    {
        $this->ensureDefaultSheet();

        if (is_string($col)) {
            $col = $this->columnNameToIndex($col);
        }

        if ($row < 1 || $col < 1) {
            throw new \Exception("Row and column must be >= 1");
        }

        $rows = $this->xlsx->rows($this->currentSheetIndex);

        $rowIndex = $row - 1;
        $colIndex = $col - 1;

        return $rows[$rowIndex][$colIndex] ?? null;
    }

    /**
     * Convert Excel column letters (A, B, AA, AB...) to 1-based index
     */
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

    /**
     * Forward unknown methods directly to SimpleXLSX
     */
    public function __call($method, $args)
    {
        if (method_exists($this->xlsx, $method)) {
            return call_user_func_array([$this->xlsx, $method], $args);
        }

        throw new \BadMethodCallException("Method $method does not exist in Reader or SimpleXLSX");
    }


}
