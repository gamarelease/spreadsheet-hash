# Spreadsheet hash function [:uk:](README.md) [:portugal:](README.pt.md)
_Portable solutions for calculating hashes in spreadsheets_

The files in this repository implement a function that can generate pseudo-random hashes from text input.  The spreadsheet files, in particular, are an example of the function usage applied to calculating anonymized identifiers from subjects' names.

## The hash function

The hash function chosen is a simple one, the [Fowler-Noll-Vo (FNV) hash function](https://en.m.wikipedia.org/wiki/Fowler%E2%80%93Noll%E2%80%93Vo_hash_function), which could be implemented reasonably without dependencies on external computing libraries.  The FNV1a 32 bit variant used here generates 32 bit hashes in the form of hexadecimal strings of length 8 (e.g. `D58B3FA7`).

## Portability

The code and spreadsheet files were tested on **LibreOffice Calc** (LC) and **Microsoft Excel** (ME), plus **WPS Spreadsheets** (WS) and **Google Spreedsheets** (GS) when possible.  Some files have versions/formats specific for Calc or Excel, as noted below.

## The files

### Formula-based implementation

- [`ID-Generator.xlsx`](https://github.com/leogama/spreadsheet-hash/releases/download/v1.0.0/Gerador-de-ID.xlsx) (1.8 MB) — uses formulas to implement the FNV1a hash (tested on LC, ME, WS and GS)
- [`Gerador-de-ID.xlsx`](https://github.com/leogama/spreadsheet-hash/releases/download/v1.0.0/Gerador-de-ID.xlsx) (1.8 MB) — identical to the file above, but with instructions in Brazilian Portuguese

The formula-based implementation doesn't depend on macros and, therefore, doesn't require special permissions to run and has better portability.  You can use it even on Google Spreedsheets!  However, it depends on a rigid workbook structure with multiple sheets.

### Macro-based implementation

- [`ID-Generator-MACRO.ods`](https://github.com/leogama/spreadsheet-hash/releases/download/v1.0.0/ID-Generator-MACRO.ods) (44.7 KB) — uses user-defined functions to implement the FNV1a hash (LibreOffice Calc format)
- [`ID-Generator-MACRO.xlsm`](https://github.com/leogama/spreadsheet-hash/releases/download/v1.0.0/ID-Generator-MACRO.xlsm) (48.7 KB) — identical to the file above, but in the Microsoft Excel format
- [`name2id.Calc.bas`](name2id.Calc.bas) — code of the BASIC/VBA module used in the macro-based spreadsheets
- [`name2id.Excel.bas`](name2id.Excel.bas) — identical to the file above, but with the file encoding and format required by Excel

The macro-based implementation is more flexible in usage, but is less portable in the sense that LibreOffice Calc can't edit `.XLSM` Excel files (althought it can open and even run them), and vice-versa.  So, even though the macro code itself is portable, there's no compatible file format.

> _Note:_ files with macros won't run on Google Spreadsheets and their execution needs to be authorized in the spreadsheet program.

### Implementations in other programming languages

Modules that implement the functions `HASH/FNV1a_32` and `NAME2ID` from `name2id.*.bas` with identical APIs and output formats.

- [`name2id.py`](name2id.py) — depends on the package `fnvhash` from PyPI
- [`name2id.R`](name2id.R) — depends on the package `bitops` from CRAN
