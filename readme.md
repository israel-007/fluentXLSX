# FluentXLSX

A fluent, developer-friendly PHP library that unifies the functionalities of [SimpleXLSX](https://github.com/shuchkin/simplexlsx) (for reading Excel files) and [SimpleXLSXGen](https://github.com/shuchkin/simplexlsxgen) (for writing Excel files).

With FluentXLSX, you can read, transform, and write Excel files in a single, consistent, and expressive API — without worrying about the differences between underlying libraries.

## Features (Planned)

* Unified API for reading and writing XLSX files.
* Fluent chaining for transformations (filter, map, take, skip).
* Multi-format conversion (XLSX ⇄ CSV ⇄ JSON ⇄ Array).
* Multiple sheets support (create, select, merge).
* Validation rules for data quality checks.
* Export helpers (save, download, stream).
* More Enhancements to come...

## Dependencies

This library is built on top of:

* [SimpleXLSX](https://github.com/shuchkin/simplexlsx) - lightweight reader for .xlsx files.
* [SimpleXLSXGen](https://github.com/shuchkin/simplexlsxgen) - lightweight generator for .xlsx files.

Both are stable, battle-tested libraries that handle the low-level complexity of the XLSX format.
FluentXLSX focuses on providing a clean, expressive, and fluent API over them.

## Installation (coming soon)

Coming soon...


## Quick Examples

### Read an Excel file

```php

$rows = XLSX::load('users.xlsx')
    ->sheet(1)
    ->get();

```

### Filter and Export
```PHP

XLSX::load('users.xlsx')
    ->sheet(1)
    ->filter(fn($row) => $row[2] === 'Active')
    ->map(fn($row) => [$row[0], strtoupper($row[1])])
    ->toXLSX('active_users.xlsx');

```

### Create a New File with Multiple Sheets
```PHP

XLSX::create()
    ->sheet('Students', [
        ['ID', 'Name', 'Grade'],
        [1, 'Alice', 'A'],
        [2, 'Bob', 'B'],
    ])
    ->sheet('Teachers', [
        ['ID', 'Name', 'Subject'],
        [1, 'Mr. Smith', 'Math'],
        [2, 'Ms. Johnson', 'Science'],
    ])
    ->save('school.xlsx');

```

## Contribution

PRs and issues are welcome! Please open a GitHub issue to discuss new ideas, bugs, or feature requests.












