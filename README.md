#  SmartXLS Excel Library for .NET

## About
[ SmartXLS](https://www.smartxls.com) is a high performance .NET component which can write, read, calculate Excel compatible files without the need for Microsoft Excel on either the developer or client machines. 
It was entirely written in 100% managed C# code.
The Excel library is used for creating, reading and manipulating MS Excel files, including support for advanced features like formatting, formulas, charts, macros, images and pivot tables.  
SmartXLS library  is designed to be easy to use, with a straightforward API and documentation. 

##  SmartXLS for .Net provides the following features

* [Import data into excel Worksheet from DataTable.](https://smartxls.com/csharp/data.htm#vdata-import-datagrid)
* [Export excel worksheet data to DataTable.](https://smartxls.com/csharp/data.htm#vdata-export-datagrid)
* [CSV files](https://smartxls.com/csharp/workbook.htm#vworkbook-rw-xls) (delimited with comma, tab, semicolon or any other separator).
* [XLSX/XLSM reading/writing](https://smartxls.com/csharp/workbook.htm#vworkbook-rw-xlsx) (Excel2007-Excel2016 openxml format).
* [Encrypte/Decrypt Excel files(xls and xlsx).](https://smartxls.com/csharp/workbook.htm#vworkbook-encrypt-decrypt-xlsx)
* Template support (create new workbooks using existing workbook as a template).
* [Multiple worksheets per file.](https://smartxls.com/csharp/worksheets.htm#vsheets-add-remove-sheets)
* Various cell data types (numbers, strings,dates, floating point etc.).
* [Custom number formatting.](https://smartxls.com/csharp/formatting.htm#vformat-number)
* [Merged regions.](https://smartxls.com/csharp/formatting.htm#vformat-merge-cells)
* Cell styles (alignment, indentation, rotation, borders, shading, protection, text wrapping and shrinking etc.).
* Font formatting (size, color, font type, italic and strikeout properties, different levels of boldness, underlining, subscript and superscript).
* formatting options([font](https://smartxls.com/csharp/formatting.htm#vformat-font),color,content format,pattern,[border](https://smartxls.com/csharp/formatting.htm#vformat-border) line,border color,[align](https://smartxls.com/csharp/formatting.htm#vformat-alignment)).
* [Copy/Move/Delete](https://smartxls.com/csharp/data.htm#vdata-range-manipulation)(Cells can be copied, moved and deleted).
* [Row height and column width.](https://smartxls.com/csharp/worksheets.htm#vsheets-rowheight-columnwidth)
* [Formula support](https://smartxls.com/csharp/data.htm#vdata-formulas) (absolute and relative references, names, 3D cell references, more than 260 supported functions).
* [Named ranges](https://smartxls.com/csharp/data.htm#vdata-named-range)(Names can be used where you want. Access cells through their name is easy).
* Supports Formula Calculations using robust Formula Calculation (Cells and whole worksheets can be calculated).
* [Formulas with references to external workbooks](https://smartxls.com/csharp/data.htm#vdata-formulas-external-workbook) are supported as well.
* Set Print Options([print area](https://smartxls.com/csharp/printing.htm#vprint-sheet),[print margins](https://smartxls.com/csharp/printing.htm#vprint-margins),[print header/footer](https://smartxls.com/csharp/printing.htm#vprint-header-footer),[page break](https://smartxls.com/csharp/printing.htm#vprint-page-breaks) etc).
* [Rich Text](https://smartxls.com/csharp/formatting.htm#vformat-richtext-formatting): Insert rich text into cells.
* [Conditional formatting](https://smartxls.com/csharp/formatting.htm#vformat-conditional-formatting).
* Data validation for cell values.
* Copy cell ranges between worksheets/workbooks..
* [Comments](https://smartxls.com/csharp/drawings-charts.htm#vdrawings-charts-add-comments): Insert cell comments.
* [Charts](https://smartxls.com/csharp/drawings-charts.htm#vdrawings-charts-chart)(Supports all standard Chart Types like Column, Bar, Line, Pie, Scatter etc,Customize Charts by setting their different properties).
* Create Excel [pivot table](https://smartxls.com/csharp/pivot-table.htm#vpivot-table) from scratch.
* Macro support: Read macros and preserve them on re-saving.
* File properties with details about the author and company that generated the Excel file or custom properties
* Written in 100% managed code.
* Unicode support.
* Supporting asp. Net Medium trust level.
* Mono support (works on Unix/Linux/OsX machines with Mono).

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


---
[Product Page](https://www.smartxls.com) / [Samples](https://smartxls.com/sample-list.htm) / [Support](https://www.smartxls.com/contact.htm)
