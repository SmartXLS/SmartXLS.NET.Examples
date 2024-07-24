using SmartXLS;
using System;
using System.Data;

namespace Examples_cs
{
    internal class DataValidationRead
    {
        static void Main(string[] args)
        {

            try
            {
                WorkBook book = new WorkBook();

                book.readXLSX("..\\..\\..\\template\\DVTemplate.xlsx");

                DataValidation validation = book.getValidation(0, 0);

                //Reading the Data Validation list
                string lists = validation.Formula1;

                lists = lists.Replace("\0", " ");

                Console.WriteLine(lists);
                Console.ReadKey();
            }
            catch (Exception e1)
            {
                System.Console.WriteLine(e1.StackTrace);
            }
        }
    }
}
