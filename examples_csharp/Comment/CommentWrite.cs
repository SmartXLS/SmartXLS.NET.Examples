using SmartXLS;
using System;

namespace Examples_cs
{
    internal class CommentWrite
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            //add a comment to B2
            workBook.addComment(1, 1, "comment text here!", "author name here!");

            workBook.writeXLSX("./Sample.xlsx");
            System.Diagnostics.Process.Start("Sample.xlsx");


        }
    }
}
