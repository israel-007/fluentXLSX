# FluentXLSX

A fluent, developer-friendly PHP library that unifies the functionalities of [SimpleXLSX](https://github.com/shuchkin/simplexlsx) (for reading Excel files) and [SimpleXLSXGen](https://github.com/shuchkin/simplexlsxgen) (for writing Excel files).

With FluentXLSX, you can read, transform, and write Excel files in a single, consistent, and expressive API — without worrying about the differences between underlying libraries.

## Features

- Unified API for reading and writing XLSX files.
- Read from file or raw data.
- Select sheets by index, name, or active sheet.
- Access rows, headers, and individual cells (by Excel reference or index).
- Write new Excel files with multiple sheets.
- Save or download generated Excel files.
- Convert sheets to associative arrays (headers as keys).

## Dependencies

This library is built on top of the following libraries which are provided by [shuchkin](https://github.com/shuchkin):

- [SimpleXLSX](https://github.com/shuchkin/simplexlsx) - lightweight reader for .xlsx files.
- [SimpleXLSXGen](https://github.com/shuchkin/simplexlsxgen) - lightweight generator for .xlsx files.

**Credit:**  
Special thanks to [Sergey Shuchkin](https://github.com/shuchkin) for creating and maintaining [SimpleXLSX](https://github.com/shuchkin/simplexlsx) and [SimpleXLSXGen](https://github.com/shuchkin/simplexlsxgen), which make this project possible.

Both are stable, battle-tested libraries that handle the low-level complexity of the XLSX format.  
FluentXLSX focuses on providing a clean, expressive, and fluent API over them.

## Installation

Install via Composer:

```sh
composer require fluentxlsx/fluentxlsx
```

## Usage Examples

### Read an Excel file

```php
use Fluentxlsx\Excel;

// Read from file, select sheet by index, get all rows
$reader = Excel::read('users.xlsx')
    ->sheet(0);

$rows = $reader->rows();
```

### Read from raw data

```php
use Fluentxlsx\Excel;

$data = file_get_contents('users.xlsx');
$reader = Excel::read($data);

$headers = $reader->headers();
```

### Select sheet by name or active sheet

```php
use Fluentxlsx\Excel;

// By name
$reader = Excel::read('users.xlsx')->sheet('Sheet1');

// By active sheet
$reader = Excel::read('users.xlsx')->sheet('ACTIVE');
```

### Get specific rows or range

```php
// First 5 rows
$rows = $reader->rows(5);

// Rows 2 to 10 (inclusive, 1-based)
$rows = $reader->rows([2, 10]);
```

### Get explicit rows

```php
// Get rows 1, 3, and 5
$rows = $reader->rowsEx([1, 3, 5]);
```

### Get headers

```php
$headers = $reader->headers();
```

### Convert to associative array

```php
$assoc = $reader->toAssoc();
// [
//   ['id' => 1, 'name' => 'Alice'],
//   ['id' => 2, 'name' => 'Bob'],
// ]
```

### Get all sheet names

```php
$sheets = $reader->sheets();
```

### Get a cell by reference or index

```php
// By Excel reference
$value = $reader->cell('B3');

// By row and column (1-based)
$value = $reader->cellEx(3, 2); // Row 3, Column 2 (B)
$value = $reader->cellEx(3, 'B');
```

## Direct Access to SimpleXLSX Methods

You can also use methods from the underlying SimpleXLSX library directly on the reader object returned by `Excel::read()`.  
This allows you to access advanced features or methods not wrapped by FluentXLSX.

**Example:**

```php
use Fluentxlsx\Excel;

// Get the dimension (rows, columns) of the first sheet
$reader = Excel::read('users.xlsx');
list($rows, $cols) = $reader->dimension();
```

Other SimpleXLSX methods can be used in the same way.  
Refer to the [SimpleXLSX documentation](https://github.com/shuchkin/simplexlsx#examples) for more details.

## Contribution

PRs and issues are welcome! Please open a GitHub issue to discuss new ideas, bugs, or feature requests.