using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes.Export
{
    internal class StatTable : ExcelTable
    {
        public StatTable(object[,] data) : base(data)
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
            double chartWidth = excelResources.CellWidth * 5;
            double chartHeight = excelResources.CellHeight * 8;
            double leftPos = excelResources.CellWidth * 7;
            double topPos = (verticalPosition - 1) * excelResources.CellHeight;

            // Create chart
            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(leftPos, topPos, chartWidth, chartHeight);
            Chart chart = chartObject.Chart;

            // Get relative data range
            Range range = GetRange(
                sheet,
                verticalPosition + 1,
                horizontalPosition + 3,
                verticalPosition + 3,
                horizontalPosition + 3
                );

            // Write chart
            chart.ChartWizard(
                range,
                XlChartType.xlColumnClustered,
                Title: _data[0, 0].ToString(),
                ValueTitle: "Percents");

            // Set chart legend
            Series series = chart.SeriesCollection(1) as Series;
            series.HasDataLabels = true;
            chart.HasLegend = false;
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);

            // Get relatice phases range
            range = GetRange(
                sheet, 
                verticalPosition + 1, 
                1, 
                verticalPosition + 3,
                1);
            xAxis.CategoryNames = range;
        }

    }
}
