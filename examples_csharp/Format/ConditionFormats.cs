using SmartXLS;
using System;

namespace Examples_cs
{
    internal class ConditionFormats
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            ConditionFormat[] condfmt = new ConditionFormat[3];
            condfmt[0] = workBook.CreateConditionFormat();
            condfmt[1] = workBook.CreateConditionFormat();
            condfmt[2] = workBook.CreateConditionFormat();

            // Condition #1
            RangeStyle cf = condfmt[0].RangeStyle;
            condfmt[0].Type = ConditionFormat.eTypeFormula;
            condfmt[0].setFormula1("and(iseven(row()), $D1 > 1000)", 0, 0);
            cf.FontColor = 0x00ff00;
            cf.Pattern = RangeStyle.PatternSolid;
            cf.PatternFG = 0xcc99ff;
            condfmt[0].RangeStyle = cf;

            // Condition #2
            condfmt[1].Type = ConditionFormat.eTypeFormula;
            condfmt[1].setFormula1("iseven($A1)", 0, 0);
            cf.FontColor = 0xffffff;
            condfmt[1].RangeStyle = cf;

            // Condition #3
            condfmt[2].Type = ConditionFormat.eTypeCell;
            condfmt[2].setFormula1("500", 0, 0);
            condfmt[2].Operator = ConditionFormat.eOperatorGreaterThan;
            cf = condfmt[2].RangeStyle;
            cf.FontColor = 0xff0000;
            condfmt[2].RangeStyle = cf;

            // Select the range and apply conditional formatting
            workBook.setSelection(0, 0, 39, 3);
            workBook.ConditionalFormats = condfmt;

            workBook.setNumber(1, 0, 1);
            workBook.setText(1, 5, "iseven($A1) no fill");
            workBook.setNumber(3, 3, 2000);
            workBook.setText(3, 5, "and(iseven(row()), $D1 > 1000) green");
            workBook.setNumber(5, 0, 601);
            workBook.setText(5, 5, "> 500 red");

            workBook.writeXLSX("./Sample.xlsx");
            System.Diagnostics.Process.Start("Sample.xlsx");

        }
    }
}
