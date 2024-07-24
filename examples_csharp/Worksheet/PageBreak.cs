using SmartXLS;
using System;

namespace Examples_cs
{
    internal class PageBreak
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();
            try
            {
                workBook.readXLSX("..\\..\\..\\template\\book.xlsx");
                workBook.copyRange(8, 9, 14, 13, 1, 1, 7, 5);
                workBook.addRowPageBreak(2);
                workBook.addColPageBreak(2);
                workBook.writeXLSX("Sample.xlsx");
                System.Diagnostics.Process.Start("Sample.xlsx");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
            }
        }
    }
}
