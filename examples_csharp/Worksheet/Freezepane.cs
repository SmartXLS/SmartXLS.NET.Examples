using SmartXLS;
using System;

namespace Examples_cs
{
    internal class Freezepane
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();
            try
            {
                workBook.read("..\\..\\..\\template\\book.xls");
                workBook.copyRange(8, 9, 14, 13, 1, 1, 7, 5);
                workBook.freezePanes(0, 0, 8, 6, false);
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
