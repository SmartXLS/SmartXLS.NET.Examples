#  SmartXLS Excel Library for .NET

## About
[ SmartXLS](https://www.smartxls.com) is a high performance .NET component which can write, read, calculate Excel compatible files without the need for Microsoft Excel on either the developer or client machines. 
It was entirely written in 100% managed C# code.
The Excel library is used for creating, reading and manipulating MS Excel files, including support for advanced features like formatting, formulas, charts, macros, images and pivot tables.  
SmartXLS library  is designed to be easy to use, with a straightforward API and documentation. 

##  SmartXLS for .Net provides the following features

* [Import data into excel Worksheet from DataTable.](https://smartxls.com/csharp/data.htm#vdata-import-datagrid)
* Export excel worksheet data to DataTable.
* CSV files (delimited with comma, tab, semicolon or any other separator).
* XLSX/XLSM reading / writing (Excel2007-Excel2016 openxml format).
* Encrypte/Decrypt Excel files(xls and xlsx).
* Template support (create new workbooks using existing workbook as a template).
* Multiple worksheets per file.
* Various cell data types (numbers, strings,dates, floating point etc.).
* Custom number formatting.
* Merged regions.
* Cell styles (alignment, indentation, rotation, borders, shading, protection, text wrapping and shrinking etc.).
* Font formatting (size, color, font type, italic and strikeout properties, different levels of boldness, underlining, subscript and superscript).
* formatting options(font,color,content format,pattern,border line,border color,align).
* Copy/Move/Delete(Cells can be copied, moved and deleted).
* Row height and column width.
* Formula support (absolute and relative references, names, 3D cell references, more than 260 supported functions).
* Named ranges(Names can be used where you want. Access cells through their name is easy).
* Supports Formula Calculations using robust Formula Calculation (Cells and whole worksheets can be calculated).
* Formulas with references to external workbooks are supported as well.
* Set Print Options(print area,print margins,print header/footer,page break etc).
* Rich Text: Insert rich text into cells.
* Conditional formatting.
* Comments: Insert cell comments.
* Charts(Supports all standard Chart Types like Column, Bar, Line, Pie, Scatter etc,Customize Charts by setting their different properties).
* Create Excel pivot table from scratch.
* Macro support: Read macros and preserve them on re-saving.
* Written in 100% managed code.
* Unicode support.
* Supporting asp. Net Medium trust level.
* Mono support (works on Unix/Linux/OsX machines with Mono).

* [Create Excel files](https://www.smartxls.com/manual/basics/create-excel-file.html), new files or from Excel templates
* [Import Excel data](https://www.smartxls.com/manual/basics/import-from-xlsx-file-format.html), modify Excel file and resave the file
* [Convert Excel files](https://www.smartxls.com/manual/basics/convert-html-to-excel.html), between MS Excel file formats (XLSX, XLSM, XLSB, XLS and SpreadsheetML) and also text formats (HTML, XML, CSV and TXT).
* [Format cells](https://www.smartxls.com/manual/basics/format-excel-cells.html), rows, and columns with background, foreground, fonts, borders, alignments, number and date formats and other formatting elements. Conditional formatting is also supported.
* Multiple sheets 
* Complex formulas and functions, named ranges and formulas, formula calculation engine included
* Hyperlinks, comments and images
* Data validation for cell values, including drop-down selection
* Print options and page breaks
* Group rows and columns, split and freeze panes, filter and auto-filter
* Charts with various supported types and formatting
* Pivot tables and pivot charts
* Encryption and password protection to protect the Excel file from unauthorized access, [protect sheet data](https://www.smartxls.com/manual/basics/excel-protect-sheet.html) inside sheet from altering
* VB code and macros preservation
* File properties with details about the author and company that generated the Excel file or custom properties
* Import/export from data structures, SQL databases, lists of data, export DataTable to Excel, import Excel to DataTable, import/export from GridView or DataGridView, import/export ResultSet to Excel

## Supported File Formats
**MS Excel Open XML:** XLSX, XLSM  
**MS Excel Binary:** XLSB, XLS  
**XML:** SpreadsheetML, XML specific schema  
**Text:** CSV, TXT  

## Getting Started in .NET

### **Get Started with SmartXLS**: Download and install SmartXLS nuget package  

Download  SmartXLS from [nuget.org](https://www.nuget.org/packages/SXPackage) and execute below line in Package Manager Console from Visual Studio:  
```Install-Package SXPackage```  
or search for  smartxls in NuGet Package Manager in Visual Studio and install.


### **Create XLSX Excel File from Scratch**: Start coding

You can execute the code below in C# to create an Excel file having two sheets and a value set in "A1" cell.

```
// Create an instance of the WorkBook class
WorkBook workbook = new WorkBook();

// Add data in A1 cell
workbook.setText(0, 0, "Hello world!");

// Create Excel file
workbook.writeXLSX("out\\Excel.xlsx");
```

## Documentation
 SmartXLS website provides detailed information on how to use the various features and functionalities of the  SmartXLS library, including a complete [User Guide](https://www.smartxls.com/manual), [tutorials](https://www.smartxls.com/manual/tutorials/ SmartXLS-tutorials.html), [demos](https://www.smartxls.com/net-excel-library#demo), and [API documentation](https://www.smartxls.com/manual/API_Documentation/index.html).

---
[Product Page](https://www.smartxls.com) / [Trial License](https://www.smartxls.com/trials) / [Getting Started](https://www.smartxls.com/manual/getting-started/welcome.html) / [Tutorials](https://www.smartxls.com/tutorials) / [Documentation](https://www.smartxls.com/manual) / [FAQ](https://www.smartxls.com/faq) / [Support](https://www.smartxls.com/ask-a-question)
