<?php

require __DIR__ . '/vendor/autoload.php';

use Fluentxlsx\Excel;

Excel::write()
    ->sheet('Report')
    ->download('report.xlsx');