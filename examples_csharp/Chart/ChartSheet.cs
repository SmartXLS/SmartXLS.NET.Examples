﻿using SmartXLS;
using System;
using System.Drawing;

namespace Examples_cs
{
    internal class ChartSheet
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();
            try
            {
                //set data
                workBook.setText(0, 1, "Jan");
                workBook.setText(0, 2, "Feb");
                workBook.setText(0, 3, "Mar");
                workBook.setText(0, 4, "Apr");
                workBook.setText(0, 5, "Jun");

                workBook.setText(1, 0, "Comfrey");
                workBook.setText(2, 0, "Bananas");
                workBook.setText(3, 0, "Papaya");
                workBook.setText(4, 0, "Mango");
                workBook.setText(5, 0, "Lilikoi");
                for (int col = 1; col <= 5; col++)
                    for (int row = 1; row <= 5; row++)
                        workBook.setFormula(row, col, "RAND()");
                workBook.setText(6, 0, "Total");
                workBook.setFormula(6, 1, "SUM(B2:B6)");
                workBook.setSelection("B7:F7");
                //auto fill the range with the first cell's formula or data
                workBook.editCopyRight();

                workBook.insertSheets(0, 1);
                workBook.Sheet = 0;
                workBook.setSheetName(0, "ChartSheet");
                ChartShape chart = workBook.addChartSheet(0);

                chart.ChartType = ChartShape.Column;
                //link data source, link each series to columns(true to rows).
                chart.setLinkRange("Sheet1!$a$1:$F$5", false);

                //set axis title
                chart.setAxisTitle(ChartShape.XAxis, 0, "X-axis data");
                chart.setAxisTitle(ChartShape.YAxis, 0, "Y-axis data");
                //set series name
                chart.setSeriesName(0, "My Series number 1");
                chart.setSeriesName(1, "My Series number 2");
                chart.setSeriesName(2, "My Series number 3");
                chart.setSeriesName(3, "My Series number 4");
                chart.setSeriesName(4, "My Series number 5");
                chart.Title = "My Chart";

                //set plot area's color to darkgray
                ChartFormat chartFormat = chart.PlotFormat;
                chartFormat.FillSolid = true;
                chartFormat.ForeColor = Color.DarkGray.ToArgb();
                chart.PlotFormat = chartFormat;

                //set series 0's color to blue
                ChartFormat seriesformat = chart.getSeriesFormat(0);
                seriesformat.FillSolid = true;
                seriesformat.ForeColor = Color.Blue.ToArgb();
                chart.setSeriesFormat(0, seriesformat);

                //set series 1's color to red
                seriesformat = chart.getSeriesFormat(1);
                seriesformat.FillSolid = true;
                seriesformat.ForeColor = Color.Red.ToArgb();
                chart.setSeriesFormat(1, seriesformat);

                //set chart title's font property
                ChartFormat titleformat = chart.TitleFormat;
                titleformat.FontSize = 14 * 20;
                titleformat.FontUnderline = true;
                chart.TitleFormat = titleformat;

                workBook.writeXLSX("./Sample.xlsx");
                System.Diagnostics.Process.Start("Sample.xlsx");
            }
            catch (System.Exception ex)
            {
                Console.Error.WriteLine(ex);
            }

        }
    }
}
