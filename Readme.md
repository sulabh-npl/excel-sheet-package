# Excel File Package

Excel file package is a composer package used to export data from database to excel sheet with proper heading and formatting of data. 

*This is more than 85% Faster than existing Solutions*

## Installation

Use the package manager [composer](https://getcomposer.org/) to install Excel File Package.

```bash
composer require sulabh/excel-file-package
```

## Usage

```php
require 'vendor/autoload.php';

use Sulabh\ExcelFilePackage\ExcelFile;
use App\Model\User;

// data to be exported, supported format is Laravel's Model.
// You can filter the data as per your need using where conditions;
$data = User::query();

// Number of data you want in a Single Sheet. 
// 1 Excel Sheet can store upto 1,048,576 rows, 
// so this value should not exceed 1,048,575. 
// 1 is for header
$chunk_size = 100000;

// List of headers in each column respectively
$header = ['S.No', 'Name', 'Email', 'Phone Number'];

// No need to use .xlsx, even if entered its filtered by the package
$file_name = "UsersExport";

// Define how the data in export should be displayed
$row_formatter = function($row) {
  return [
    'SERIAL_NO', // This is a wildcard and represents an incremental value, It starts with 1 in every sheet
    $row->first_name." ".$row->last_name,
    $row->email,
    $row->phone_number,
  ];
};

// Optional
$total_count = $data->count();

ExcelFile::createExcelFile($data, $chunk_size, $header, $file_name, $row_formatter, $total_data);
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first
to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License

[MIT](https://choosealicense.com/licenses/mit/)
