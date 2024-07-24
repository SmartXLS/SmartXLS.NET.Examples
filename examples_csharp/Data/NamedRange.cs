using SmartXLS;
using System;
using System.Data;

namespace Examples_cs
{
    internal class NamedRange
    {
        static void Main(string[] args)
        {

            try
            {
                WorkBook workBook = new WorkBook();
                workBook.read("..\\..\\..\\template\\NamedRanges.xls");

                workBook.setDefinedName("Products", "$A$1:$A$6");

                workBook.setDefinedName("One", "$C$3");
                workBook.setDefinedName("Two", "$D$3");
                workBook.setSelection("E3");
                workBook.setFormula(2, 4, "SUM(One, Two)");
                workBook.recalc();

                workBook.writeXLSX("Sample.xlsx");
                System.Diagnostics.Process.Start("Sample.xlsx");
            }
            catch (Exception e1)
            {
                System.Console.WriteLine(e1.StackTrace);
            }
        }
    }
}
