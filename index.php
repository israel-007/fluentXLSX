<?php

require __DIR__ . '/vendor/autoload.php';

use Fluentxlsx\Excel;

$rows = Excel::read('users.xlsx')->toAssoc();

print_r($rows);