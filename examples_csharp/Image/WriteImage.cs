using SmartXLS;
using System;

namespace Examples_cs
{
    internal class WriteImage
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            try
            {

                //Inserting image
                workBook.addPicture(1, 0, 3, 8, "..\\..\\..\\template\\image1.png");

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
