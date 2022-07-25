using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes.Export
{
    internal class DuplicatesTable : ExcelTable
    {
        public DuplicatesTable(object[,] data) : base(data)
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
            // Set chart dimensions and positions
            ExcelResources excelResources = ExcelResources.GetInstance();
            double chartWidth = excelResources.CellWidth * 17;
            double chartHeight = excelResources.CellHeight * 15;
            double leftPos = excelResources.CellWidth * 4;
            double topPos = 1;

            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(leftPos, topPos, chartWidth, chartHeight);
            Chart chart = chartObject.Chart;

            Range range = sheet.Range[$"A1:A{_data.GetLength(0)}"];
            chart.ChartWizard(range, Gallery: XlChartType.xlXYScatterLinesNoMarkers);
            chart.PlotBy = XlRowCol.xlColumns;
            chart.HasTitle = false;
            chart.HasLegend = false;
            chart.SeriesCollection(1).XValues = sheet.Range[$"A1:A{_data.GetLength(0)}"];
            chart.SeriesCollection(1).Values = sheet.Range[$"B1:B{_data.GetLength(0)}"];
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            Axis yAxis = chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            xAxis.MaximumScale = 30000;
            yAxis.MajorUnit = 1;
        }
    }
}
