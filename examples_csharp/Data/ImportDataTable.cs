using SmartXLS;
using System;
using System.Data;

namespace Examples_cs
{
    internal class ImportDataTable
    {
        static void Main(string[] args)
        {

            try
            {
                //Imports Data from the Template spreadsheet into the Grid.

                WorkBook m_book = new WorkBook();
                m_book.read("..\\..\\..\\template\\NorthwindDataTemplate.xls");

                //Read data from spreadsheet.
                DataTable customersTable = m_book.ExportDataTable();

                System.IO.TextWriter tw = new System.IO.StringWriter();
                customersTable.WriteXml(tw, true);
                customersTable.WriteXmlSchema(tw);
                Console.WriteLine(tw.ToString());
                Console.ReadKey();
            }
            catch (Exception e1)
            {
                System.Console.WriteLine(e1.StackTrace);
            }
        }
    }
}
