## Excel-To-Access
===============
Simple basic application that help combine excel files into a DB.

### Description of the problem:
Some external softwares or websites export excel files, however the excel files have various formats. 
While Access Databases have an import function, which can import excel files which are formatted as tables, having any global data anywhere on the file will prevent the user from importing.

### The solution
The class TableBuilder is used to read a CSV file and transfer it to the Database. Different implementations of the class can handle various formats and insert it into the database in a correct form.
The solution contains two projects:
 1. CSVtoDB - a project with the classes used to make a database from excel files. The 3 basic formats are:
  1. BasicTableBuilder, when the entire file is one big table.
  2. SimpleHeaderTable - when the entire file is one big table,except for the first row which is a header.
  3. HeaderAndAdditionalDataTable - when the data starts at some row, and before that there are some global data, which we want to describe each of the rows in the current document.
 2. ExcelToAccessExample - a simple WPF application that converts an excel file to a csv, and insert the data into a chosen access database.


### Notes:
 1. The project is half finished, but it's enough for what I need right now, so I'll probably update it later.
 2. Implementing various TableBuilders will require some sort of factory for them, which should read the excel/csv file in order to decide which builder should be used.
