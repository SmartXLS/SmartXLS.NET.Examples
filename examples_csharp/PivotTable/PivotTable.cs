using SmartXLS;
using System;

namespace Examples_cs
{
    internal class PivotTable
    {
        static void Main(string[] args)
        {
            WorkBook workBook = new WorkBook();

            try
            {
                workBook.read("..\\..\\..\\template\\PivotTable.xls");

                BookPivotRangeModel model = workBook.getPivotModel();
                //Sets the source range that should be used for the PivotRange
                model.setList("A1:D27");
                //Defines the location of the PivotRange.
                model.setLocation(0, 17, 5);

                //make the cell active(to make the pivotRange selected)
                workBook.setActiveCell(17, 5);
                //get the currently selected PivotRange
                BookPivotRange pivotRange = model.getActivePivotRange();
                //refresh the pivot table from the data source
                model.refreshRange(pivotRange);

                //get the Area object associated with the PivotRange.
                RangeArea rangeArea = pivotRange.getRangeArea();

                Console.WriteLine("PivotRange Scope:" + rangeArea.ToString());
                //pivotRange.addFormulaField("double amount:", "=Amount*2");

                BookPivotArea rowArea = pivotRange.getArea(BookPivotRange.row);
                BookPivotArea columnArea = pivotRange.getArea(BookPivotRange.column);
                BookPivotArea dataArea = pivotRange.getArea(BookPivotRange.data);
                BookPivotArea pageArea = pivotRange.getArea(BookPivotRange.page);

                BookPivotField rowField = pivotRange.getField("Who");
                rowArea.addField(rowField);
                BookPivotField dataField = pivotRange.getField("Amount");
                //BookPivotField dataField = pivotRange.getField("double amount:");
                dataArea.addField(dataField);
                BookPivotField columnField = pivotRange.getField("What");
                columnArea.addField(columnField);
                BookPivotField pageField = pivotRange.getField("Week");
                pageArea.addField(pageField);

                workBook.writeXLSX("Sample.xlsx");
                System.Diagnostics.Process.Start("Sample.xlsx");
            }
            catch (Exception ex)
            {
                System.Console.Out.WriteLine(ex.Message);
            }
        }
    }
}
