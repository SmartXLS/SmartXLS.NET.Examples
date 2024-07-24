using SmartXLS;
using System;

namespace Examples_cs
{
    internal class ReadImage
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            workBook.readXLSX("..\\..\\..\\template\\book.xlsx");
            PictureShape picShape = workBook.getPictureShape(0);

            byte[] imagedata = picShape.PictureImageData;

            System.IO.FileStream fos = new System.IO.FileStream("pic.jpg", System.IO.FileMode.Create);
            fos.Write(imagedata, 0, imagedata.Length);
            fos.Close();

            System.Diagnostics.Process.Start("pic.jpg");

        }
    }
}
