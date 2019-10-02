# ASME Sign In Handler
*A small, minute effort to handle and organize sign in forms for the American Society of Mechanical Engineers, UCLA*

This program maintains a collection of attendance entries from sign-in forms. There are options to add new forms and output stored data, via a simple text-based interface

## Requirements:
- pandas
- xlrd and openpyxl (For Excel)
- Python 3.7

## Usage (Available commands)
**help**: Display help page (this page)
**add**: Add a new Excel Spreadsheet to the collection of records
**add from directory**: Given a directory, find all spreadsheets and add to records
**output**: Create an Excel spreadsheet using all records
**reset**: Delete all records
**quit**: quit this program