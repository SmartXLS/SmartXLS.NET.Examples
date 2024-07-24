using SmartXLS;
using System;
using System.Data;

namespace Examples_cs
{
    internal class DataValidationWrite
    {
        static void Main(string[] args)
        {

            try
            {
                WorkBook workBook = new WorkBook();

                workBook.setText(0, 1, "Apple");
                workBook.setText(0, 2, "Orange");
                workBook.setText(0, 3, "Banana");
                workBook.setText(6, 3, "the input value in cell C7 must be between 0 to 10");

                DataValidation dataValidation = workBook.CreateDataValidation();
                dataValidation.Type = DataValidation.eUser;
                dataValidation.Formula1 = "\"dddd\x0000gggg\x0000hhh\"";
                workBook.setSelection("A1:A5");
                workBook.DataValidation = dataValidation;

                dataValidation = workBook.CreateDataValidation();
                dataValidation.Type = DataValidation.eUser;
                dataValidation.Formula1 = "$B$1:$D$1";
                workBook.setSelection("B1:D5");
                workBook.DataValidation = dataValidation;

                dataValidation = workBook.CreateDataValidation();
                dataValidation.ShowErrorMessage = true;
                dataValidation.Type = DataValidation.eInteger;
                dataValidation.Operator = DataValidation.eBetween;
                dataValidation.Formula1 = "0";
                dataValidation.Formula2 = "10";
                workBook.setSelection(6, 2, 6, 2);//select C7
                workBook.DataValidation = dataValidation;

                workBook.writeXLSX(".\\Sample.xlsx");
                System.Diagnostics.Process.Start("Sample.xlsx");

            }
            catch (Exception e1)
            {
                System.Console.WriteLine(e1.StackTrace);
            }
        }
    }
}
