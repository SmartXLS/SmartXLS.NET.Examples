using SmartXLS;
using System;

namespace Examples_cs
{
    internal class RichTextRead
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            workBook.readXLSX("..\\..\\..\\template\\book.xlsx");

            string rft = workBook.getRichText(0, 0);

            Console.WriteLine(rft);
            Console.ReadKey();

        }
    }
}
