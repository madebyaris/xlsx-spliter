# XLSX Splitter

A Node.js application that efficiently splits large Excel (XLSX) files into smaller chunks while maintaining the original structure and formatting. This tool is particularly useful when dealing with Excel files that are too large to handle in memory.

I tested using 980K rows and it can split in 3-4 minutes for 100K rows per file, and without any issues.

If you wanted to use it on Node server, I suggest to use xlsx-stream-reader and try to load/stream via file, using in-memory buffer may lead an issue if the file is too large. 

## Features

- Split large XLSX files into smaller chunks
- Maintain headers in each split file
- Stream processing for memory efficiency
- Configurable number of rows per output file
- Preserves data types and formatting
- Progress tracking during processing
- Handles files with multiple hundred thousand rows

## Prerequisites

Before you begin, ensure you have the following installed:
- Node.js (v16.0.0 or higher)
- npm (Node Package Manager)

## Installation

1. Clone or download this repository:
```bash
git clone https://github.com/madebyaris/xlsx-spliter.git
# or download and extract the ZIP file
```

2. Navigate to the project directory:
```bash
cd xlsx-splitter
```

3. Install dependencies:
```bash
npm install
```

This will install the required packages:
- xlsx: For Excel file handling
- xlsx-stream-reader: For streaming large Excel files
- yargs: For command-line argument parsing

## Usage

The basic command syntax is:

```bash
node index.js --input <path-to-xlsx-file> --rows <rows-per-file>
```

### Parameters:

- `--input` or `-i`: (Required) Path to the input XLSX file
- `--rows` or `-r`: (Optional) Number of rows per split file (default: 1000)

### Examples:

1. Split a file with default settings (1000 rows per file):
```bash
node index.js --input ./path/to/large-file.xlsx
```

2. Split a file into chunks of 100,000 rows each:
```bash
node index.js --input ./path/to/large-file.xlsx --rows 100000
```

### Output

- Split files will be saved in the `output` directory
- Files will be named using the pattern: `original-filename_part1.xlsx`, `original-filename_part2.xlsx`, etc.
- Each file will include the original headers
- Progress will be displayed in the console during processing

## Notes

1. The first row of the input file is assumed to be the header row
2. Empty rows are skipped during processing
3. The output directory will be created automatically if it doesn't exist
4. All split files will maintain the original column structure and data types

## Troubleshooting

1. If you get memory errors:
   - Try reducing the number of rows per file
   - Ensure you have enough free disk space

2. If the process seems slow:
   - This is normal for very large files as they are processed streaming
   - Progress updates will show in the console

3. If files appear empty or corrupted:
   - Check if the input file is a valid XLSX file
   - Ensure you have write permissions in the output directory

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
