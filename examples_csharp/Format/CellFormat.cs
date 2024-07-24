using SmartXLS;
using System;
using System.Drawing;

namespace Examples_cs
{
    internal class CellFormat
    {
        static void Main(string[] args)
        {
            new CellFormat().Format();
        }


        int StartRow, StartCol, EndRow, EndCol;
        string StartRange, eRange, hdrRange, colRange, ftrRange, bodyRange;
        WorkBook workBook;
        RangeStyle m_RangeStyle;

        void Format()
        {

            workBook = new WorkBook();
            workBook.NumSheets = 12;
            workBook.setSheetName(0, "Simple Format");
            workBook.setSheetName(1, "Classic1 Format");
            workBook.setSheetName(2, "Classic3 Format");
            workBook.setSheetName(3, "Accounting1 Format");
            workBook.setSheetName(4, "Accounting2 Format");
            workBook.setSheetName(5, "Accounting3 Format");
            workBook.setSheetName(6, "Effects3D1 Format");
            workBook.setSheetName(7, "Colorful1 Format");
            workBook.setSheetName(8, "Colorful2 Format");
            workBook.setSheetName(9, "Colorful3 Format");
            workBook.setSheetName(10, "List1 Format");
            workBook.setSheetName(11, "List3 Format");

            workBook.Sheet = 0;
            setData();
            simpleFormat();

            workBook.Sheet = 1;
            setData();
            Classic1();

            workBook.Sheet = 2;
            setData();
            Classic3();

            workBook.Sheet = 3;
            setData();
            Accounting1();

            workBook.Sheet = 4;
            setData();
            Accounting2();

            workBook.Sheet = 5;
            setData();
            Accounting3();

            workBook.Sheet = 6;
            setData();
            Effects3D1();

            workBook.Sheet = 7;
            setData();
            Colorful1();

            workBook.Sheet = 8;
            setData();
            Colorful2();

            workBook.Sheet = 9;
            setData();
            Colorful3();

            workBook.Sheet = 10;
            setData();
            List1();

            workBook.Sheet = 11;
            setData();
            List3();

            workBook.writeXLSX(".\\Sample.xlsx");
            System.Diagnostics.Process.Start("Sample.xlsx");
        }
        private void setData()
        {
            workBook.setText(1, 2, "Jan");
            workBook.setText(1, 3, "Feb");
            workBook.setText(1, 4, "Mar");
            workBook.setText(1, 5, "Apr");
            workBook.setText(2, 1, "Bananas");
            workBook.setText(3, 1, "Papaya");
            workBook.setText(4, 1, "Mango");
            workBook.setText(5, 1, "Lilikoi");
            workBook.setText(6, 1, "Comfrey");
            workBook.setText(7, 1, "Total");
            workBook.setFormula(2, 2, "RAND()*100");
            workBook.setSelection(2, 2, 2, 5);
            workBook.editCopyRight();
            workBook.setSelection(2, 2, 6, 5);
            workBook.editCopyDown();
            workBook.setFormula(7, 2, "SUM(C3:C7)");
            workBook.setSelection("C8:F8");
            workBook.editCopyRight();

            StartRow = 1;
            StartCol = 1;
            EndRow = 7;
            EndCol = 5;
            StartRange = workBook.formatRCNr(StartRow, StartCol, false);
            eRange = workBook.formatRCNr(StartRow, EndCol, false);
            hdrRange = StartRange + ":" + eRange;

            eRange = workBook.formatRCNr(EndRow, StartCol, false);
            colRange = StartRange + ":" + eRange;

            StartRange = workBook.formatRCNr(EndRow, StartCol, false);
            eRange = workBook.formatRCNr(EndRow, EndCol, false);
            ftrRange = StartRange + ":" + eRange;

            StartRange = workBook.formatRCNr(StartRow + 1, StartCol + 1, false);
            eRange = workBook.formatRCNr(EndRow - 1, EndCol, false);
            bodyRange = StartRange + ":" + eRange;

            m_RangeStyle = workBook.getRangeStyle();
        }

        private void simpleFormat()
        {
            workBook.setSelection(colRange);
            AdjustFont(Color.Black.ToArgb(), true, false, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            workBook.setRangeStyle(m_RangeStyle, EndRow, StartCol, EndRow, EndCol);

            workBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentRight;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            workBook.setRangeStyle(m_RangeStyle, EndRow, StartCol, EndRow, EndCol);
        }

        private void Classic1()
        {
            workBook.setSelection(colRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderNone;
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            m_RangeStyle.RightBorder = RangeStyle.BorderThin;
            AdjustFont(Color.Black.ToArgb(), true, false, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(hdrRange);
            m_RangeStyle.VerticalInsideBorder = RangeStyle.BorderNone;
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            AdjustFont(System.Drawing.Color.Black.ToArgb(), false, true, false);
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentRight;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Classic3()
        {
            workBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            AdjustFont(Color.White.ToArgb(), true, true, false);
            AlignRight();
            SetSolidPattern(workBook.getPaletteEntry(11), Color.Black.ToArgb());
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            SetSolidPattern(workBook.getPaletteEntry(15), Color.Black.ToArgb());
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(bodyRange);
            SetSolidPattern(workBook.getPaletteEntry(15), Color.Black.ToArgb());
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(colRange);
            SetSolidPattern(workBook.getPaletteEntry(15), Color.Black.ToArgb());
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Accounting1()
        {
            workBook.setSelection(hdrRange);
            System.String numberFormat;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            AdjustFont(Color.Magenta.ToArgb(), true, true, false);
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(bodyRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderDouble;
            m_RangeStyle.CustomFormat = numberFormat;
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Accounting2()
        {
            System.String numberFormat;

            workBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderThick;
            m_RangeStyle.TopBorderColor = Color.LightGray.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorderColor = Color.LightGray.ToArgb();
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(bodyRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderThick;
            m_RangeStyle.BottomBorderColor = Color.LightGray.ToArgb();
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.TopBorderColor = Color.LightGray.ToArgb();
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Accounting3()
        {
            System.String numberFormat;
            workBook.setSelection(colRange);
            AdjustFont(Color.Black.ToArgb(), false, true, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(bodyRange);
            numberFormat = "#,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(bodyRange);
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.TopBorder = RangeStyle.BorderNone;
            m_RangeStyle.BottomBorder = RangeStyle.BorderNone;
            m_RangeStyle.CustomFormat = numberFormat;
            workBook.setRangeStyle(m_RangeStyle);

            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(bodyRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorder = RangeStyle.BorderDouble;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(bodyRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderNone;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = Color.Green.ToArgb();
            AdjustFont(workBook.getPaletteEntry(16), false, true, false);
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Effects3D1()
        {
            workBook.setSelection(1, 1, 7, 5);
            SetSolidPattern(Color.LightGray.ToArgb(), 0);

            Set3DBorder(2, 2, 6, 5, Color.LightGray.ToArgb(), Color.DarkGray.ToArgb(), Color.LightGray.ToArgb());
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(hdrRange);
            AdjustFont(Color.Magenta.ToArgb(), true, false, false);
            AlignCenter();
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(colRange);
            Set3DBorder(1, 1, 7, 1, Color.LightGray.ToArgb(), Color.LightGray.ToArgb(), Color.DarkGray.ToArgb());
            AdjustFont(Color.Black.ToArgb(), true, false, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            Set3DBorder(7, 1, 7, 5, Color.LightGray.ToArgb(), Color.DarkGray.ToArgb(), Color.DarkGray.ToArgb());
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Colorful1()
        {
            workBook.setSelection(1, 1, 7, 5);
            int color = Color.Red.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            m_RangeStyle.BottomBorderColor = Color.Red.ToArgb();
            SetSolidPattern(Color.DarkGray.ToArgb(), Color.Black.ToArgb());

            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.LeftBorder = RangeStyle.BorderMedium;
            m_RangeStyle.RightBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = Color.Red.ToArgb();
            m_RangeStyle.LeftBorderColor = Color.Red.ToArgb();
            m_RangeStyle.RightBorderColor = Color.Red.ToArgb();

            AdjustFont(color, false, false, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(hdrRange);
            SetSolidPattern(Color.Black.ToArgb(), Color.Black.ToArgb());
            AdjustFont(color, true, true, false);
            AlignCenter();
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(colRange);
            SetSolidPattern(workBook.getPaletteEntry(11), Color.Black.ToArgb());
            AdjustFont(color, true, true, false);
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Colorful2()
        {
            int color = workBook.getPaletteEntry(14);
            workBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderThin;
            SetSolidPattern(workBook.getPaletteEntry(9), Color.Black.ToArgb());
            AdjustFont(color, true, true, false);
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(colRange);
            AdjustFont(Color.Black.ToArgb(), true, true, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(1, 1, 7, 5);
            SetHatchPattern(workBook.getPaletteEntry(16), Color.Red.ToArgb());
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Colorful3()
        {
            workBook.setSelection(1, 1, 7, 5);
            SetSolidPattern(Color.Black.ToArgb(), Color.Black.ToArgb());
            AdjustFont(Color.White.ToArgb(), false, false, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(hdrRange);
            AdjustFont(Color.Green.ToArgb(), true, true, false);
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(colRange);
            AdjustFont(Color.Magenta.ToArgb(), true, true, false);
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void List1()
        {
            workBook.setSelection(1, 1, 7, 5);
            System.String newSelection, numberFormat;
            int fcolor, bcolor;
            m_RangeStyle.TopBorder = RangeStyle.BorderThin;
            m_RangeStyle.LeftBorder = RangeStyle.BorderThin;
            m_RangeStyle.RightBorder = RangeStyle.BorderThin;
            m_RangeStyle.TopBorderColor = workBook.getPaletteEntry(16);
            m_RangeStyle.LeftBorderColor = workBook.getPaletteEntry(16);
            m_RangeStyle.RightBorderColor = workBook.getPaletteEntry(16);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(hdrRange);
            SetSolidPattern(workBook.getPaletteEntry(14), Color.Black.ToArgb());
            AdjustFont(Color.Blue.ToArgb(), true, true, false);
            AlignCenter();
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            SetSolidPattern(workBook.getPaletteEntry(14), Color.Black.ToArgb());
            AdjustFont(Color.Blue.ToArgb(), true, false, false);
            numberFormat = "$ #,##0.00_);(#,##0.00)";
            m_RangeStyle.CustomFormat = numberFormat;
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);

            fcolor = workBook.getPaletteEntry(19);
            bcolor = Color.Red.ToArgb();
            for (int i = StartRow + 1; i <= EndRow - 1; i = i + 2)
            {
                newSelection = workBook.formatRCNr(i, StartCol, false) + ":" + workBook.formatRCNr(i, EndCol, false);
                workBook.setSelection(newSelection);
                SetHatchPattern(fcolor, bcolor);
                workBook.setRangeStyle(m_RangeStyle);
            }

            fcolor = Color.White.ToArgb();
            bcolor = workBook.getPaletteEntry(15);
            for (int i = StartRow + 2; i < EndRow - 1; i = i + 2)
            {
                newSelection = workBook.formatRCNr(i, StartCol, false) + ":" + workBook.formatRCNr(i, EndCol, false);
                workBook.setSelection(newSelection);
                SetHatchPattern(fcolor, bcolor);
                workBook.setRangeStyle(m_RangeStyle);
            }
        }

        private void List3()
        {
            workBook.setSelection(colRange);
            workBook.setRangeStyle(m_RangeStyle);
            workBook.setSelection(hdrRange);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = Color.DarkGray.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = Color.DarkGray.ToArgb();
            AlignCenter();
            AdjustFont(workBook.getPaletteEntry(11), true, false, false);
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(ftrRange);
            m_RangeStyle = workBook.getRangeStyle(EndRow, StartCol, EndRow, EndCol);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = Color.DarkGray.ToArgb();
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = Color.DarkGray.ToArgb();
            AlignRight();
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void Set3DBorder(int row1, int col1, int row2, int col2, int outlineColor, int rightColor, int bottomColor)
        {
            workBook.setSelection(row1, col1, row2, col2);
            m_RangeStyle.TopBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.LeftBorder = RangeStyle.BorderMedium;
            m_RangeStyle.RightBorder = RangeStyle.BorderMedium;
            m_RangeStyle.TopBorderColor = outlineColor;
            m_RangeStyle.BottomBorderColor = outlineColor;
            m_RangeStyle.LeftBorderColor = outlineColor;
            m_RangeStyle.RightBorderColor = outlineColor;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(row1, col2, row2, col2);
            m_RangeStyle.RightBorder = RangeStyle.BorderMedium;
            m_RangeStyle.RightBorderColor = rightColor;
            workBook.setRangeStyle(m_RangeStyle);

            workBook.setSelection(row2, col1, row2, col2);
            m_RangeStyle.BottomBorder = RangeStyle.BorderMedium;
            m_RangeStyle.BottomBorderColor = bottomColor;
            workBook.setRangeStyle(m_RangeStyle);
        }

        private void SetHatchPattern(int fcolor, int bcolor)
        {
            m_RangeStyle.Pattern = (short)4;
            m_RangeStyle.PatternFG = fcolor;
            m_RangeStyle.PatternBG = bcolor;
        }

        private void AlignCenter()
        {
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentCenter;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            m_RangeStyle.WordWrap = false;
        }

        private void AdjustFont(int color, bool bold, bool italic, bool underline)
        {
            m_RangeStyle.FontBold = bold;
            m_RangeStyle.FontItalic = italic;
            m_RangeStyle.FontUnderline = RangeStyle.UnderlineSingle;
            m_RangeStyle.FontColor = color;
        }

        private void AlignRight()
        {
            m_RangeStyle.HorizontalAlignment = RangeStyle.HorizontalAlignmentRight;
            m_RangeStyle.VerticalAlignment = RangeStyle.VerticalAlignmentBottom;
            m_RangeStyle.WordWrap = false;
        }

        private void SetSolidPattern(int fcolor, int bcolor)
        {
            short nPattern;
            nPattern = 1;
            m_RangeStyle.Pattern = nPattern;
            m_RangeStyle.PatternFG = fcolor;
            m_RangeStyle.PatternBG = bcolor;
        }

    }
}
