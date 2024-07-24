using SmartXLS;
using System;

namespace Examples_cs
{
    internal class RowHeightColWidth
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();
            try
            {
                workBook.readXLSX("..\\..\\..\\template\\book.xlsx");

                workBook.setRowHeight(1, 25 * 1440 / 256);
                workBook.setColWidth(1, 25 * 256);

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
