using SmartXLS;
using System;

namespace Examples_cs
{
    internal class ChartEdit
    {
        static void Main(string[] args)
        {
            try
            {
                WorkBook workBook = new WorkBook();

                //read in
                workBook.readXLSX("..\\..\\..\\template\\chartTemplate.xlsx");

                //get chartshape from sheet 1
                ChartShape chartShape = workBook.getChart(0);

                chartShape.ChartType = ChartShape.Bar;
                chartShape.Title = "Chart 1";
                //change 3D chart to 2D
                chartShape.set3Dimensional(false);

                //select sheet 2
                workBook.Sheet = 1;
                //get chartshape in the sheet
                chartShape = workBook.getChart(0);
                //change chart type to step
                chartShape.ChartType = ChartShape.Step;
                //set axis title
                chartShape.setAxisTitle(ChartShape.XAxis, 0, "X-axis data");
                chartShape.setAxisTitle(ChartShape.YAxis, 0, "Y-axis data");
                chartShape.Title = "Chart 2";
                //change chart to 3D
                chartShape.set3Dimensional(true);

                workBook.writeXLSX("./Sample.xlsx");
                System.Diagnostics.Process.Start("Sample.xlsx");
            }
            catch (Exception ex)
            {
                System.Console.Out.WriteLine(ex);
            }
        }
    }
}
