using SmartXLS;
using System;

namespace Examples_cs
{
    internal class HyperlinkWrite
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            //add a url link to F6
            workBook.addHyperlink(5, 5, 5, 5, "http://www.smartxls.com/", HyperLink.kURLAbs, "Hello,web url hyperlink!");

            //add a file link to F7
            workBook.addHyperlink(6, 5, 6, 5, "c:\\", HyperLink.kFileAbs, "file link");

            workBook.writeXLSX("Sample.xlsx");
            System.Diagnostics.Process.Start("Sample.xlsx");

        }
    }
}
