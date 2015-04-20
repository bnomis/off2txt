# off2txt

Extracts ASCII/Unicode text from Office files to separate files.

Useful if you have a document containing two languages (e.g. English and Chinese) and you want to separate the languages into text files for further processing and analysis. 

Supports Open XML file formats. That is, docx, pptx, and xlsx.

Word and PowerPoint files are extracted to text files. 
Excel files are extracted to CSV files, columns are preserved.

Can be used to make a CSV file from Excel without opening Excel.

## Examples

### Extract ASCII and Unicode Text From a Word Document

```shell
$ off2txt -s word.docx
```

The above will make two files: word-ascii.txt and word-unicode.txt

### Extract ASCII and Unicode Text From an Excel Document

```shell
$ off2txt -s excel.xlsx
```

The above will make two files: excel-ascii.csv and excel-unicode.csv

## Notes

If an extracted file would be empty, it is not created.

Excel is different. Columns are preserved. So may get a CSV file of empty columns. Cells are put in the extracted ASCII file if they containt ASCII only otherwise they are streamed to the Unicode file.


## Usage

```shell
usage: off2txt [options] File [File ...]

off2txt: extract ASCII/Unicode text from Office files to separate files

positional arguments:
  File                  Files to extract from

optional arguments:
  -h, --help            show this help message and exit
  --version             show program's version number and exit
  --debug               Turn on debug logging.
  --debug-log FILE      Save debug logging to FILE.
  -a EXTENSION, --ascii EXTENSION
                        Identifier to append to input file name to make ASCII
                        output file name when splitting Unicode and ASCII
                        text. Default ascii.
  -d DIRECTORY, --directory DIRECTORY
                        Save extracted text to DIRECTORY. Ignored if the -o
                        option is given.
  -e EXTENSION, --extension EXTENSION
                        Extension to use for extracted text files. Default for
                        Word and PowerPoint is txt. Default for Excel is csv.
  -o FILE, --output FILE
                        Save extracted text to FILE. If not given, the output
                        file is named the same as the input file but with a
                        txt extension. The extension can be changed with the
                        -e option. Files are opened in append mode unless the
                        -X option is given.
  -s, --split           Split ASCII and Unicode text into two separate files.
                        Unicode files are named by adding -unicode before the
                        file extension. The Unicode identifer can be changed
                        with the -u option.
  -u EXTENSION, --unicode EXTENSION
                        Identifier to append to input file name to make
                        Unicode output file name when splitting Unicode and
                        ASCII text. Default unicode.
  -A, --suppress-file-access-errors
                        Do not print file/directory access errors.
  -X, --overwrite-output-files
                        Truncate output files before writing.
```

                         