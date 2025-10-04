<?php

namespace Fluentxlsx;

class Excel
{
    /**
     * Entry point for reading
     *
     * @param string|null $source  File path or raw data
     * @return Reader
     * @throws \Exception
     */
    public static function read($source = null)
    {
        $reader = new Reader();

        if ($source !== null) {
            if (is_string($source) && file_exists($source)) {
                $reader->fromFile($source);
            } else {
                $reader->fromData($source);
            }
        }

        return $reader;
    }

    /**
     * Entry point for writing
     *
     * @return Writer
     */
    public static function write()
    {
        return new Writer();
    }
}
