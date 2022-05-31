using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes.Export
{
    internal class GraphTableWithChart : ExcelTable
    {
        public GraphTableWithChart(object[,] data) : base(data)
        {
        }

        public override int ExportToSheet(_Worksheet sheet, int verticalPosition, int horizontalPosition)
        {
            WriteData(sheet, verticalPosition, horizontalPosition);
            Decorate(sheet, verticalPosition, horizontalPosition);
            WriteChart(sheet, verticalPosition, horizontalPosition);
            return verticalPosition + _data.GetLength(0);
        }

        private void WriteChart(_Worksheet sheet, int verticalPosition, int horizontalPosition)
        {
            ExcelResources excelResources = ExcelResources.GetInstance();
            double chartWidth = excelResources.CellWidth * 10;
            double chartHeight = excelResources.CellHeight * 15;
            double leftPos = 
                ((horizontalPosition - 1) * excelResources.CellWidth) + 
                (excelResources.CellWidth * _data.GetLength(1)) + 
                (2 * excelResources.CellWidth);
            double topPos = 1;

            // Get range to determine width of chart
            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(leftPos, topPos, chartWidth, chartHeight);
            Chart chart = chartObject.Chart;

            Range range = GetRange(
                sheet,
                verticalPosition + 1,
                horizontalPosition,
                verticalPosition + 3,
                horizontalPosition + _data.GetLength(1) - 1);
            chart.ChartWizard(
                range,
                XlChartType.xlColumnClustered,
                PlotBy: XlRowCol.xlRows,
                Title: _data[0, 0].ToString(),
                ValueTitle: "Percents");
            chart.HasLegend = true;
            chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            range = GetRange(
                sheet, 
                verticalPosition, 
                horizontalPosition + 1, 
                verticalPosition, 
                horizontalPosition + _data.GetLength(1) - 1);
            xAxis.CategoryNames = range;
        }
    }
}
