# Excel Sheet Protection Remover

## Description

This Python script allows you to remove sheet protection from Excel files stored in a zip archive. The script processes the Excel file, removes the sheet protection from worksheets, and saves the modified contents in a new Excel archive.

## Installation

1. Clone the repository or download the source code files.

2. Install the required libraries. The script uses `argparse` for handling command-line arguments. You can install them using pip:

```shell
pip install argparse
```

## Usage

To run the script, use the following command-line format:

```python
python main.py path/to/your/file.xslx
```

Replace `path/to/your/file.xlsx` with the path to the Excel file you want to process.

Example:

```python
python main.py data/file_example_XLSX_10.xlsx.
```
**Note:** There is a sample Excel file named ```file_example_XLSX_10.xlsx```available in the ```data``` directory. You can use this file to test the script.

The script will process the specified Excel file, remove sheet protection from the worksheets, and save the modified contents in a new Excel file with the prefix "modified_". The new Excel file will be created in the same directory as the original Excel file.


## Notes
This script will only remove sheet protection from Excel files that are not password protected. If a file requires a password to open or modify, the script will not be able to remove the sheet protection.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributions

Contributions to this project are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## Disclaimer

This script is intended for educational and informational purposes only. Use it responsibly and only on files that you have the right to modify. The author and contributors are not responsible for any misuse or damage resulting from the use of this script. Always make backups of your files before processing them with this or any other script.
