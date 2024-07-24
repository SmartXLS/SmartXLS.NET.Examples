using SmartXLS;
using System;
using System.Drawing;

namespace Examples_cs
{
    internal class RichTextWrite
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            //set data
            workBook.setText(0, 0, "Hello, you are welcome!");

            //text orientation
            RangeStyle rangeStyle = workBook.getRangeStyle();
            rangeStyle.Orientation = (short)45;
            workBook.setRangeStyle(rangeStyle);

            //multi text selection format
            workBook.setTextSelection(0, 6);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.FontBold = true;
            workBook.setRangeStyle(rangeStyle);

            workBook.setTextSelection(7, 10);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.FontItalic = true;
            rangeStyle.FontColor = Color.IndianRed.ToArgb();
            workBook.setRangeStyle(rangeStyle);

            workBook.setTextSelection(11, 14);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.FontUnderline = RangeStyle.UnderlineSingle;
            workBook.setRangeStyle(rangeStyle);

            workBook.setTextSelection(15, 22);
            rangeStyle = workBook.getRangeStyle();
            rangeStyle.FontSize = 14 * 20;
            workBook.setRangeStyle(rangeStyle);

            workBook.writeXLSX("Sample.xlsx");
            System.Diagnostics.Process.Start("Sample.xlsx");
        }
    }
}
