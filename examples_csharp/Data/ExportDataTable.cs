using SmartXLS;
using System;
using System.Data;

namespace Examples_cs
{
    internal class ExportDataTable
    {
        static void Main(string[] args)
        {

            try
            {
                //Load Data
                DataSet customersDataSet = new DataSet();
                customersDataSet.ReadXml("..\\..\\..\\template\\Customers.xml", XmlReadMode.ReadSchema);
                DataTable northwindDt = customersDataSet.Tables[0];

                WorkBook workBook = new WorkBook();

                workBook.ImportDataTable(northwindDt, true, 0, 0, -1, -1);

                workBook.writeXLSX("./Sample.xlsx");
                System.Diagnostics.Process.Start("Sample.xlsx");

            }
            catch (Exception e1)
            {
                System.Console.WriteLine(e1.StackTrace);
            }
        }
    }
}
