# iternova/docxmerger - Merge DOCX files

A simple library to merge DOCX files. 

More info: ITERNOVA [https://www.iternova.net]

## Requirements

- PHP 7.0 or higher.

## Installation

To install easily using composer (and packagist.org), complete your `composer.json` file with:

```javascript
{
    "require": {
        "iternova/docxmerger": "dev-master"
    }
}
```


## Usage

```php

$absolute_path_directory = __DIR__.'/../../tmp/';
$array_input_files = [
    $absolute_path_directory . 'input_file_01.docx',
    $absolute_path_directory . 'input_file_02.docx',
    $absolute_path_directory . 'input_file_03.docx',
];
$array_page_breaks = [
    false,
    true,
    false,
];
$output_file = $absolute_path_directory .'output_file.docx';

$docx_merger = new \Iternova\DOCXMerger\DOCXMerger();
$docx_merger->add_files( $array_input_files, $array_page_breaks );
$docx_merger->save( $output_file);

```

## License

This package is released under the MIT license. You are free to use, modify and distribute this software or any variant of it.
