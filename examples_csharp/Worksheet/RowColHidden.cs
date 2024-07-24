using SmartXLS;
using System;

namespace Examples_cs
{
    internal class RowColHidden
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();
            try
            {
                workBook.readXLSX("..\\..\\..\\template\\book.xlsx");
                workBook.setRowHidden(1, true);
                workBook.setColHidden(1, true);
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
