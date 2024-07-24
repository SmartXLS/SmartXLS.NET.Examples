using SmartXLS;
using System;

namespace Examples_cs
{
    internal class HyperlinkRead
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            workBook.readXLSX("..\\..\\..\\template\\book.xlsx");

            // get the first index from the current sheet
            HyperLink hyperLink = workBook.getHyperlink(0);
            string txtHyperlink = hyperLink.LinkString;
            Console.WriteLine(txtHyperlink);
            Console.ReadKey();
        }
    }
}
