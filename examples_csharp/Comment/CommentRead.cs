using SmartXLS;
using System;

namespace Examples_cs
{
    internal class CommentRead
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            workBook.readXLSX("..\\..\\..\\template\\book.xlsx");

            // get the first index from the current sheet
            CommentShape commentShape = workBook.getComment(1, 7);
            if (commentShape != null)
            {
                //string text = "Author:" + commentShape.Author + "\n text:" + commentShape.Text;
                string text = commentShape.RichText;

                Console.WriteLine(text);
                Console.ReadKey();
            }


        }
    }
}
