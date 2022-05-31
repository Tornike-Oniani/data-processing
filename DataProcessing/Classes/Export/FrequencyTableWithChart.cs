using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes.Export
{
    internal class FrequencyTableWithChart : ExcelTable
    {
        public FrequencyTableWithChart(object[,] data) : base(data)
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

            // chartTop = WriteStatChart(statsSheet, tableInfo, 5, 10, tableCount, chartTop);
            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(leftPos, topPos, chartWidth, chartHeight);
            Chart chart = chartObject.Chart;

            // This is weird but it works only if this table is on position 1,1 (we might want to scale this to accomodate state change)
            Range range = GetRange(
                sheet,
                2,
                1,
                _data.GetLength(0),
                _data.GetLength(1));
            chart.ChartWizard(
                range,
                XlChartType.xlColumnClustered,
                Title: _data[0,0],
                ValueTitle: "Frequency");
            foreach (Series series in chart.SeriesCollection())
            {
                series.HasDataLabels = true;
            }
            chart.HasLegend = true;
            chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;
            //Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            //range = GetRange(customFrequencySheet, 2, 1, 2, tableInfo.Table.Columns.Count - 2);
            //xAxis.CategoryNames = range;
        }
    }
}
