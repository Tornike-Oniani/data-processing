using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DataProcessing.Models;
using DataProcessing.Repositories;
using DataProcessing.Utils;
using Microsoft.Office.Interop.Excel;

namespace DataProcessing.Classes
{
    enum DataType
    {
        Second,
        Minute,
        Percentage,
        Number
    }

    class ExcelManager
    {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private Color primaryDark = Color.FromArgb(75, 177, 250);
        private Color secondaryDark = Color.FromArgb(250, 148, 75);
        private Color primaryLight = Color.FromArgb(148, 216, 255);
        private Color secondaryLight = Color.FromArgb(255, 187, 148);
        private Color timeMarkColor = Color.FromArgb(202, 255, 138);
        private Color darkLightMarkColor = Color.FromArgb(250, 228, 102);
        private Color alternateColor = Color.FromArgb(230, 229, 225);
        private Color errorColor = Color.FromArgb(232, 128, 128);
        private Color criteriaLight = Color.FromArgb(255, 157, 148);
        private Color criteriaDark = Color.FromArgb(250, 92, 75);
        private Workfile workfile = WorkfileManager.GetInstance().SelectedWorkFile;
        private int distanceBetweenTables = 2;
        private ExportOptions options;
        private List<DataTableInfo> statTableCollection;
        private List<DataTableInfo> graphTableCollection;
        private Services services = Services.GetInstance();
        private List<int> markerLocations = new List<int>();
        private List<int> darkLightMarkerLocations = new List<int>();
        private double cellWidth;
        private double cellHeight;
        private double chartLeft;
        private double chartLeftAlt;
        private double chartVerticalDistance;
        private double graphLeft;
        private double duplicatedChartLeft;
        List<int> statTablePositions = new List<int>();

        public ExcelManager(ExportOptions options, List<DataTableInfo> statTableCollection, List<DataTableInfo> graphTableCollection)
        {
            this.options = options;
            this.statTableCollection = statTableCollection;
            this.graphTableCollection = graphTableCollection;
        }

        public async Task<List<int>> CheckExcelFile(string filePath)
        {
            services.SetWorkStatus(true);

            List<int> errorRows = new List<int>();

            await Task.Run(() =>
            {
                // 1. Open excel
                services.UpdateWorkStatus("Opening excel...");
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(filePath, ReadOnly: false);
                Worksheet worksheet = workbook.Sheets[1];

                Range firstRow = worksheet.Cells[1, 1];
                if (firstRow.Value == null)
                {
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(workbook);
                    excel.Workbooks.Close();
                    excel.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    throw new Exception("First row can not be blank in excel file!");
                }

                // 2. Count nonblank rows
                services.UpdateWorkStatus("Counting rows...");
                Range targetCells = worksheet.UsedRange;
                Range allCells = targetCells.Cells;
                Range allRows = targetCells.Rows;
                object[,] allValues = (object[,])allCells.Value;
                int rowCount = 0;

                for (int i = 1; i <= allRows.Count; i++)
                {
                    if (allValues[i, 1] == null) { break; }
                    rowCount++;
                }

                // Create array of raw data
                Range rangeData = worksheet.Range[$"A1:B{rowCount}"];
                Range rangeCells = rangeData.Cells;
                object[,] dataValues = rangeCells.Value;
                List<Tuple<TimeSpan, int>> timeData = new List<Tuple<TimeSpan, int>>();
                for (int i = 1; i <= dataValues.Length / 2; i++)
                {
                    string sTimeSpan = DateTime.FromOADate((double)dataValues[i, 1]).ToString("HH:mm:ss");
                    int[] nTimeData = new int[3] {
                        int.Parse(sTimeSpan.Split(':')[0]),
                        int.Parse(sTimeSpan.Split(':')[1]),
                        int.Parse(sTimeSpan.Split(':')[2])
                    };
                    string val = "";
                    if (dataValues[i, 2] != null)
                        val = dataValues[i, 2].ToString().Trim();
                    int thisState = dataValues[i, 2] == null ? 0 : int.Parse(dataValues[i, 2].ToString().Trim());
                    TimeSpan span = new TimeSpan(nTimeData[0], nTimeData[1], nTimeData[2]);

                    Tuple<TimeSpan, int> tuple = new Tuple<TimeSpan, int>(span, thisState);
                    timeData.Add(tuple);
                }

                // Check time integrity
                for (int i = 1; i < timeData.Count - 2; i++)
                {
                    if (!isBetweenTimeInterval(timeData[i - 1].Item1, timeData[i + 1].Item1, timeData[i].Item1))
                    {
                        errorRows.Add(i + 1);
                    }
                }

                // Supposed to clean excel from memory but fails miserably
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                excel.Workbooks.Close();
                excel.Quit();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GetExcelProcess(excel).Kill();
            });

            services.SetWorkStatus(false);

            return errorRows;
        }
        public async Task HighlightExcelFileErrors(string filePath, List<int> errorRows)
        {
            services.SetWorkStatus(true);

            await Task.Run(() =>
            {
                // 1. Open excel
                services.UpdateWorkStatus("Opening excel...");
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(filePath, ReadOnly: false);
                Worksheet worksheet = workbook.Sheets[1];

                Range errorRange;
                Interior interior;
                foreach (int errorRow in errorRows)
                {
                    errorRange = worksheet.Range[$"A{errorRow}:B{errorRow}"];
                    interior = errorRange.Interior;
                    interior.Color = errorColor;
                }

                excel.Visible = true;
                excel.UserControl = true;
            });

            services.SetWorkStatus(false);
        }
        public async Task ImportFromExcel(string filePath)
        {
            services.SetWorkStatus(true);

            await Task.Run(() =>
            {
                // 1. Open excel
                services.UpdateWorkStatus("Starting import...");
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(filePath, ReadOnly: false);
                Worksheet worksheet = workbook.Sheets[1];

                if (worksheet.Cells[1, 1].Value == null)
                {
                    Marshal.ReleaseComObject(worksheet);
                    Marshal.ReleaseComObject(workbook);
                    excel.Workbooks.Close();
                    excel.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GetExcelProcess(excel).Kill();
                    throw new Exception("First row can not be blank in excel file!");
                }

                // 2. Count nonblank rows
                services.UpdateWorkStatus("Counting rows...");
                Range targetRange = worksheet.UsedRange;
                Range targetCells = targetRange.Cells;
                Range allRows = targetCells.Rows;
                object[,] allValues = (object[,])targetCells.Value;
                int rowCount = 0;

                for (int i = 1; i <= allRows.Count; i++)
                {
                    if (allValues[i, 1] == null) { break; }
                    rowCount++;
                }

                // Create array of raw data
                targetRange = worksheet.Range[$"A1:B{rowCount}"];
                targetCells = targetRange.Cells;
                object[,] dataValues = (object[,])targetCells.Value;
                List<Tuple<TimeSpan, int>> timeData = new List<Tuple<TimeSpan, int>>();
                for (int i = 1; i <= dataValues.Length / 2; i++)
                {
                    string sTimeSpan = DateTime.FromOADate((double)dataValues[i, 1]).ToString("HH:mm:ss");
                    int[] nTimeData = new int[3] {
                        int.Parse(sTimeSpan.Split(':')[0]),
                        int.Parse(sTimeSpan.Split(':')[1]),
                        int.Parse(sTimeSpan.Split(':')[2])
                    };
                    string val = "";
                    if (dataValues[i, 2] != null)
                        val = dataValues[i, 2].ToString().Trim();
                    int thisState = dataValues[i, 2] == null ? 0 : int.Parse(dataValues[i, 2].ToString().Trim());
                    TimeSpan span = new TimeSpan(nTimeData[0], nTimeData[1], nTimeData[2]);

                    Tuple<TimeSpan, int> tuple = new Tuple<TimeSpan, int>(span, thisState);
                    timeData.Add(tuple);
                }

                List<TimeStamp> samples = new List<TimeStamp>();

                for (int i = 0; i < timeData.Count; i++)
                {
                    updateStatus("Importing data", i + 1, rowCount);
                    TimeSpan span = timeData[i].Item1;
                    int state = timeData[i].Item2;
                    TimeStamp sample = new TimeStamp() { Time = span, State = state };
                    samples.Add(sample);
                }

                // Supposed to clean excel from memory but fails miserably
                Process excelProcess = GetExcelProcess(excel);
                while (Marshal.ReleaseComObject(allRows) > 0);
                while (Marshal.ReleaseComObject(targetCells) > 0);
                while (Marshal.ReleaseComObject(targetRange) > 0);
                allRows = null;
                targetCells = null;
                targetRange = null;
                while (Marshal.ReleaseComObject(worksheet) > 0);
                worksheet = null;
                excel.Workbooks.Close();
                while (Marshal.ReleaseComObject(workbook) > 0) ;
                workbook = null;
                while (Marshal.ReleaseComObject(excel.Workbooks) > 0) ;
                excel.Quit();
                while (Marshal.ReleaseComObject(excel) > 0) ;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                excelProcess.Kill();

                // Add markers
                services.UpdateWorkStatus("Adding markers...");
                bool hasLightMarker = false;
                bool hasDarkMarker = false;
                TimeSpan lightMarker = new TimeSpan(10, 0, 0);
                TimeSpan darkMarker = new TimeSpan(22, 0, 0);
                List<TimeStamp> markAddedSamples = new List<TimeStamp>();
                if (samples[0].Time == lightMarker) { hasLightMarker = true; }
                if (samples[1].Time == darkMarker) { hasDarkMarker = true; }
                markAddedSamples.Add(samples[0]);
                for (int i = 1; i < samples.Count; i++)
                {
                    TimeStamp curSample = samples[i];
                    TimeStamp prevSample = samples[i - 1];
                    if (curSample.Time == lightMarker) { hasLightMarker = true; curSample.IsMarker = true; }
                    if (curSample.Time == darkMarker) { hasDarkMarker = true; curSample.IsMarker = true; }

                    if (!hasLightMarker && isBetweenTimeInterval(prevSample.Time, curSample.Time, lightMarker))
                    {
                        markAddedSamples.Add(new TimeStamp() { Time = lightMarker, State = curSample.State, IsMarker = true });
                    }

                    if (!hasDarkMarker && isBetweenTimeInterval(prevSample.Time, curSample.Time, darkMarker))
                    {
                        markAddedSamples.Add(new TimeStamp() { Time = darkMarker, State = curSample.State, IsMarker = true });
                    }

                    markAddedSamples.Add(curSample);
                }

                // Persist to database
                Services.GetInstance().UpdateWorkStatus($"Persisting data");
                TimeStamp.SaveMany(markAddedSamples);
            });

            services.SetWorkStatus(false);
        }
        public async Task ExportToExcel(List<TimeStamp> timeStamps, List<Tuple<int, int>> duplicatedTimes, List<int> hourRowIndexes)
        {
            services.SetWorkStatus(true);
            await Task.Run(() =>
            {
                // 1. Open excel
                Application excel = new Application();
                excel.Caption = WorkfileManager.GetInstance().SelectedWorkFile.Name;
                _Workbook wb = excel.Workbooks.Add(Missing.Value);
                _Worksheet rawDataSheet = wb.ActiveSheet;
                rawDataSheet.Name = "Raw Data";

                // Raw Data
                services.UpdateWorkStatus("Exporting raw data");
                WriteRawDataList(rawDataSheet, timeStamps, 1);

                Range formatRange = rawDataSheet.Range["A:B"];
                formatRange.NumberFormat = "[h]:mm:ss";
                formatRange = rawDataSheet.Range["C:C"];
                formatRange.NumberFormat = "General";
                formatRange.NumberFormat = "General";

                int prev = 1;
                formatRange = rawDataSheet.Range[$"A{prev}:E{prev}"];
                formatRange.Interior.Color = timeMarkColor;
                for (int i = 0; i < hourRowIndexes.Count; i++)
                {
                    formatRange = rawDataSheet.Range[$"A{hourRowIndexes[i]}:E{hourRowIndexes[i]}"];
                    formatRange.Interior.Color = timeMarkColor;
                    formatRange = rawDataSheet.Range[$"F{prev}:F{hourRowIndexes[i]}"];
                    formatRange.Merge();
                    formatRange.Value = $"Hour {(i + 1) * options.TimeMark}";
                    formatRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    formatRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    if (i % 2 == 0) formatRange.Interior.Color = alternateColor;
                    prev = hourRowIndexes[i] + 1;
                }
                //for (int i = 0; i < markerLocations.Count; i++)
                //{
                //    formatRange = rawDataSheet.Range[$"A{markerLocations[i]}:E{markerLocations[i]}"];
                //    formatRange.Interior.Color = timeMarkColor;
                //}
                for (int i = 0; i < darkLightMarkerLocations.Count; i++)
                {
                    formatRange = rawDataSheet.Range[$"A{darkLightMarkerLocations[i]}:E{darkLightMarkerLocations[i]}"];
                    formatRange.Interior.Color = darkLightMarkColor;
                }

                Range[] ranges = new Range[4] { rawDataSheet.Range["A1"], rawDataSheet.Range["B1"], rawDataSheet.Range["C1"], rawDataSheet.Range["D1"] };
                foreach (Range r in ranges)
                    r.EntireColumn.AutoFit();

                // Stats table
                services.UpdateWorkStatus("Exporting stats tables");
                _Worksheet statsSheet = CreateNewSheet(wb, "Stats", 1);
                statsSheet.EnableSelection = XlEnableSelection.xlNoSelection;

                int position = 1;
                statTablePositions.Add(position);
                int index = 1;
                int max = statTableCollection.Count;
                foreach (DataTableInfo tableInfo in statTableCollection)
                {
                    updateStatus("Stat tables", index, max);
                    position = WriteDataTable(statsSheet, tableInfo, position);
                    position += distanceBetweenTables;
                    statTablePositions.Add(position);
                    index++;
                }

                statsSheet.Range["A1"].EntireColumn.AutoFit();

                // Create stat charts
                Range range = statsSheet.Range["A1:G1"];
                chartLeft = range.Width;
                range = statsSheet.Range["A1:N1"];
                chartLeftAlt = range.Width;
                range = statsSheet.Range["G1:G2"];
                chartVerticalDistance = range.Height;
                range = range = statsSheet.Range["G1:G1"];
                cellWidth = range.Width;
                cellHeight = range.Height;

                int tableCount = 1;
                double chartTop = 0;
                index = 1;
                foreach (DataTableInfo tableInfo in statTableCollection)
                {
                    updateStatus("Stat charts", index, max);
                    chartTop = WriteStatChart(statsSheet, tableInfo, 5, 10, tableCount, chartTop);
                    tableCount++;
                    index++;
                }

                // Graph table
                services.UpdateWorkStatus("Exporting graph tables");
                _Worksheet graphSheet = CreateNewSheet(wb, "Graph Stats", 2);
                graphSheet.EnableSelection = XlEnableSelection.xlNoSelection;

                position = 1;
                Range lastHour;
                index = 1;
                max = graphTableCollection.Count;
                foreach (DataTableInfo tableInfo in graphTableCollection)
                {
                    updateStatus("Graph tables", index, max);
                    position = WriteDataTable(graphSheet, tableInfo, position);
                    lastHour = graphSheet.Cells[position, tableInfo.HeaderIndexes.Item2];
                    lastHour.EntireColumn.AutoFit();
                    position += distanceBetweenTables;
                    index++;
                }

                graphSheet.Range["A1"].EntireColumn.AutoFit();

                range = graphSheet.Range["A1:K1"];
                graphLeft = range.Width;
                // Width is calculated by columns for now (parameter 10 doesn't do anything)
                WriteGraphChart(graphSheet, graphTableCollection[0], 10, 15, 1);

                // Duplicated graph
                services.UpdateWorkStatus("Exporting duplicated graph times");
                _Worksheet duplicatesSheet = CreateNewSheet(wb, "Duplicated Graph Stats", 3);
                duplicatesSheet.EnableSelection = XlEnableSelection.xlNoSelection;

                WriteDuplicatesList(duplicatesSheet, duplicatedTimes, 1);

                range = duplicatesSheet.Range["A1:D1"];
                duplicatedChartLeft = range.Width;
                WriteDuplicatesChart(duplicatesSheet, duplicatedTimes.Count, 17, 15);

                wb.Sheets[1].Select(Type.Missing);
                excel.Visible = true;
                excel.UserControl = true;
            });
            services.SetWorkStatus(false);
        }

        private int WriteDataTable(_Worksheet sheet, DataTableInfo tableInfo, int position)
        {
            int curPosition = position;
            System.Data.DataTable table = tableInfo.Table;

            object[,] values = DataTableTo2DArray(table);
            Range start = sheet.Cells[curPosition, 1];
            Range end = sheet.Cells[curPosition + table.Rows.Count, table.Columns.Count];
            Range tableRange = sheet.Range[start, end];
            tableRange.Value = values;

            // Color header
            start = sheet.Cells[position, 1];
            end = sheet.Cells[position, tableInfo.HeaderIndexes.Item2];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? secondaryDark : secondaryLight;

            // Color phases
            start = sheet.Cells[position + tableInfo.PhasesIndexes.Item1, 1];
            end = sheet.Cells[position + tableInfo.PhasesIndexes.Item2, 1];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? primaryDark : primaryLight;

            // Color criteria
            if (tableInfo.CriteriaPhases != null)
            {
                if (tableInfo.CriteriaPhases.Item2 > 0)
                {
                    start = sheet.Cells[position + tableInfo.CriteriaPhases.Item1 + 1, 1];
                    end = sheet.Cells[position + tableInfo.CriteriaPhases.Item1 + tableInfo.CriteriaPhases.Item2, 1];
                    tableRange = sheet.Range[start, end];
                    tableRange.Interior.Color = tableInfo.IsTotal ? criteriaDark : criteriaLight;
                }
            }

            // Set alignments
            start = sheet.Cells[position, 2];
            end = sheet.Cells[position, tableInfo.HeaderIndexes.Item2];
            tableRange = sheet.Range[start, end];
            tableRange.HorizontalAlignment = XlHAlign.xlHAlignRight;

            return curPosition + table.Rows.Count + 1;
        }
        private void WriteRawDataList(_Worksheet sheet, List<TimeStamp> timeStamps, int position)
        {
            object[,] values = ListTo2DArray(timeStamps);
            Range start = sheet.Cells[position, 1];
            Range end = sheet.Cells[position + timeStamps.Count - 1, 5];
            Range tableRange = sheet.Range[start, end];
            tableRange.Value = values;
        }
        private void WriteDuplicatesList(_Worksheet sheet, List<Tuple<int, int>> duplicates, int position)
        {
            object[,] values = ListTo2DArray(duplicates);
            Range start = sheet.Cells[position, 1];
            Range end = sheet.Cells[position + duplicates.Count - 1, 2];
            Range tableRange = sheet.Range[start, end];
            tableRange.Value = values;
        }
        private double WriteStatChart(_Worksheet sheet, DataTableInfo tableInfo, int width, int height, int tableCount, double chartTop)
        {
            // Create chart
            double leftPos = chartLeft;
            if (tableCount > 1)
            {
                leftPos = tableCount % 2 == 0 ? chartLeftAlt : chartLeft;
            }
            double topPos = chartTop;
            if (tableCount > 2)
            {
                topPos = (tableCount - 1) % 2 == 0 ? topPos + height * cellHeight + chartVerticalDistance : topPos;
            }

            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(leftPos, topPos, width * cellWidth, height * cellHeight);
            Chart chart = chartObject.Chart;

            int tablePos = statTablePositions[tableCount - 1];

            Range range = GetRange(sheet, tablePos + 1, 4, tablePos + options.MaxStates, 4);
            chart.ChartWizard(
                range,
                XlChartType.xlColumnClustered,
                Title: tableInfo.Table.TableName,
                ValueTitle: "Percents");
            Series series = chart.SeriesCollection(1) as Series;
            series.HasDataLabels = true;
            chart.HasLegend = false;
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            range = GetRange(sheet, tablePos + 1, 1, tablePos + 3, 1);
            xAxis.CategoryNames = range;

            return topPos;
        }
        private void WriteGraphChart(_Worksheet sheet, DataTableInfo tableInfo, int width, int height, int tablePos)
        {
            Range range = GetRange(sheet, 1, 1, 1, tableInfo.Table.Columns.Count);
            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(0, graphTableCollection.Count * 6 * cellHeight, range.Width, height * cellHeight);
            Chart chart = chartObject.Chart;

            range = GetRange(sheet, tablePos + 1, 1, tablePos + options.MaxStates, tableInfo.Table.Columns.Count);
            chart.ChartWizard(
                range,
                XlChartType.xlColumnClustered,
                Title: tableInfo.Table.TableName,
                ValueTitle: "Percents");
            chart.HasLegend = true;
            chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            range = GetRange(sheet, tablePos, 2, tablePos, tableInfo.Table.Columns.Count);
            xAxis.CategoryNames = range;
        }
        private void WriteDuplicatesChart(_Worksheet sheet, int duplicateCount, int width, int height)
        {
            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(duplicatedChartLeft, 1, width * cellWidth, height * cellHeight);
            Chart chart = chartObject.Chart;

            Range range = sheet.Range[$"A1:A{duplicateCount}"];
            chart.ChartWizard(range, Gallery: XlChartType.xlXYScatterLinesNoMarkers);
            chart.PlotBy = XlRowCol.xlColumns;
            chart.HasTitle = false;
            chart.HasLegend = false;
            chart.SeriesCollection(1).XValues = sheet.Range[$"A1:A{duplicateCount}"];
            chart.SeriesCollection(1).Values = sheet.Range[$"B1:B{duplicateCount}"];
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            Axis yAxis = chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
            xAxis.MaximumScale = 30000;
            yAxis.MajorUnit = 1;
        }

        private object[,] DataTableTo2DArray(System.Data.DataTable table)
        {
            // Rows + 1 because we also have to add column names
            object[,] table2D = new object[table.Rows.Count + 1, table.Columns.Count];

            // Title
            table2D[0, 0] = table.TableName;

            // Column names
            for (int i = 1; i < table.Columns.Count; i++)
            {
                table2D[0, i] = table.Columns[i].ColumnName;
            }

            // Table contents
            DataRow curRow;
            DataColumn curColumn;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                curRow = table.Rows[i];
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    curColumn = table.Columns[j];
                    table2D[i + 1, j] = curRow[curColumn.ColumnName];
                }
            }

            return table2D;
        }
        private object[,] ListTo2DArray(List<TimeStamp> timeStamps)
        {
            object[,] list2D = new object[timeStamps.Count, 5];

            TimeStamp curTimeStamp;
            for (int i = 0; i < timeStamps.Count; i++)
            {
                curTimeStamp = timeStamps[i];
                if (curTimeStamp.IsMarker) { darkLightMarkerLocations.Add(i + 1); }
                if (curTimeStamp.IsTimeMarked) { markerLocations.Add(i + 1); }
                list2D[i, 0] = curTimeStamp.Time.ToString();
                list2D[i, 1] = curTimeStamp.TimeDifference.ToString();
                list2D[i, 2] = curTimeStamp.TimeDifferenceInDouble;
                list2D[i, 3] = curTimeStamp.TimeDifferenceInSeconds;
                list2D[i, 4] = curTimeStamp.State;

            }

            return list2D;
        }
        private object[,] ListTo2DArray(List<Tuple<int, int>> duplicatedTimeAndStates)
        {
            object[,] list2D = new object[duplicatedTimeAndStates.Count, 2];

            Tuple<int, int> curDuplicatedTimeAndState;
            for (int i = 0; i < duplicatedTimeAndStates.Count; i++)
            {
                curDuplicatedTimeAndState = duplicatedTimeAndStates[i];
                list2D[i, 0] = curDuplicatedTimeAndState.Item1;
                list2D[i, 1] = curDuplicatedTimeAndState.Item2;
            }

            return list2D;
        }

        private _Worksheet CreateNewSheet(_Workbook workbook, string name, int position)
        {
            Sheets sheets = workbook.Sheets;
            Worksheet sheet = sheets.Add(Type.Missing, sheets[position], Type.Missing, Type.Missing);
            sheet.Name = name;
            return sheet;
        }
        private void ColorRange(_Worksheet sheet, Range start, Range end, Color color)
        {
            sheet.Range[start, end].Interior.Color = color;
        }
        private void updateStatus(string label, int index, int count)
        {
            if (index % 10 == 0 || index == count) { services.UpdateWorkStatus($"{label} {index}/{count}"); }
        }
        private bool isBetweenTimeInterval(TimeSpan from, TimeSpan till, TimeSpan time)
        {
            if (from < till)
            {
                return from <= time && time <= till;
            }
            else
            {
                return from <= time || time <= till;
            }
        }
        private static Range GetRange(_Worksheet sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Range start = sheet.Cells[startRow, startColumn];
            Range end = sheet.Cells[endRow, endColumn];
            return sheet.Range[start, end];
        }
        Process GetExcelProcess(Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        /* BACKUP
public async Task ExportToExcel(List<TimeStamp> records, string folderPath, string fileName)
{
    services.SetWorkStatus(true);
    await Task.Run(() =>
    {
        // 1. Open excel
        Application excel = new Application();
        excel.Caption = WorkfileManager.GetInstance().SelectedWorkFile.Name;
        _Workbook wb = excel.Workbooks.Add(Missing.Value);
        _Worksheet rawDataSheet = wb.ActiveSheet;
        rawDataSheet.Name = "Raw Data";
        excel.AutoRecover.Enabled = false;

        CreateAndWriteRawDataSheet(records, rawDataSheet);
        CreateAndWriteStatsSheet(wb);
        CreateAndWriteGraphSheet(wb);
        CreateAndWriteDuplicateSheet(wb);

        // Finish
        wb.Sheets[1].Select(Type.Missing);
        wb.SaveAs(Path.Combine(folderPath, fileName + ".xlsx"), XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false,
            XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
        //excel.Visible = true;
        //excel.UserControl = true;
        excel.Quit();
        Process excelProcess = GetExcelProcess(excel);
        excelProcess.Kill();
    });
    services.SetWorkStatus(false);
}
private void CreateAndWriteRawDataSheet(List<TimeStamp> records, _Worksheet rawDataSheet)
{
    ExportSamples(records, rawDataSheet);
    FormatColumns(rawDataSheet);
    ExportRawDataCalculations(workfile, rawDataSheet);
}
private void CreateAndWriteStatsSheet(_Workbook wb)
{
    _Worksheet hourlyCalcSheet = CreateNewSheet(wb, "Stats", 1);

    WriteStatsTable(hourlyCalcSheet, "Total", workfile.Stats, workfile.StatesMapping, 1, true);

    int step = workfile.StatesMapping.Count + distanceBetweenTables + additionalDistanceFromSpecificCrietria;
    int curPosition = 2 + step;
    int index = 1;
    int count = workfile.HourlyStats.Count;
    foreach (KeyValuePair<int, Stats> entry in workfile.HourlyStats)
    {
        updateStatus("Calculating hourly stats", index, count);
        string header = entry.Value.TotalTime != options.TimeMark * 3600 && index == count ? $"Last {Math.Round(workfile.LastSectionTime / 60, 2)} hour - {workfile.LastSectionTime} minutes" : $"{entry.Key * options.TimeMark} hour";
        WriteStatsTable(hourlyCalcSheet, header, entry.Value, workfile.StatesMapping, curPosition);
        curPosition += step;
        index++;
    }
    hourlyCalcSheet.Range["A1"].EntireColumn.AutoFit();
}
private void CreateAndWriteGraphSheet(_Workbook wb)
{
    _Worksheet graphCalcSheet = CreateNewSheet(wb, "Graph Stats", 2);
    int originalStep = workfile.StatesMapping.Count + distanceBetweenTables + 2;
    int step = originalStep;
    WriteGraphTable(graphCalcSheet, "Percentages %", workfile.StatesMapping, 1, DataType.Percentage);
    WriteGraphTable(graphCalcSheet, "Minutes", workfile.StatesMapping, step, DataType.Minute);
    step += originalStep - 1;
    WriteGraphTable(graphCalcSheet, "Seconds", workfile.StatesMapping, step, DataType.Second);
    step += originalStep - 1;
    WriteGraphTable(graphCalcSheet, "Numbers", workfile.StatesMapping, step, DataType.Number);
    graphCalcSheet.Range["A1"].EntireColumn.AutoFit();
}
private void CreateAndWriteDuplicateSheet(_Workbook wb)
{
    _Worksheet calcDuplicateSheet = CreateNewSheet(wb, "Duplicated Graph Stats", 3);
    int position = 1;
    int index = 1;
    int count = workfile.DuplicatedTimes.Count;
    foreach (Tuple<int, int> duplicate in workfile.DuplicatedTimes)
    {
        updateStatus("Duplicating data", index, count);
        WriteDuplicate(calcDuplicateSheet, duplicate, position);
        position++;
        index++;
    }
}
private void ExportSamples(List<TimeStamp> records, _Worksheet sheet)
{
    TimeStamp sample;
    int count = records.Count;
    for (int i = 0; i < count; i++)
    {
        updateStatus("Exporting raw data", i + 1, count);
        sample = records[i];
        sheet.Cells[i + 1, 1] = sample.Time.ToString();
        sheet.Cells[i + 1, 2] = sample.TimeDifference.ToString();
        sheet.Cells[i + 1, 3] = sample.TimeDifferenceInDouble;
        sheet.Cells[i + 1, 4] = sample.TimeDifferenceInSeconds;
        sheet.Cells[i + 1, 5] = sample.State;

        // Coloring markers
        //if (sample.IsTimeMarked) // Moved into RawDataCalculations
        //{
        //    markerLocations.Add(i + 1);
        //    sheet.Range[sheet.Cells[i + 1, 1], sheet.Cells[i + 1, 5]].Interior.Color = timeMarkColor;
        //}
        if (sample.IsMarker) 
        {
            markerLocations.Add(i + 1);
            sheet.Range[sheet.Cells[i + 1, 1], sheet.Cells[i + 1, 5]].Interior.Color = darkLightMarkColor;
        }
    }
}
private void ExportRawDataCalculations(Workfile workfile, _Worksheet rawDataSheet)
{
    Dictionary<int, int[]> indexes = workfile.HourlyIndexes;
    Range hourRange;
    Range hourMarkerRange;
    bool gray = false;
    int count = indexes.Count;
    int index = 1;
    foreach (KeyValuePair<int, int[]> entry in indexes)
    {
        updateStatus("Indexing hours", index, count);
        hourRange = rawDataSheet.Range[$"F{entry.Value[0]}:F{entry.Value[1]}"];
        hourRange.Merge();
        hourRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        hourRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        // Color timeMarked rows
        hourMarkerRange = rawDataSheet.Range[$"A{entry.Value[1]}:E{entry.Value[1]}"];
        hourMarkerRange.Interior.Color = timeMarkColor;
        if (gray) { hourRange.Interior.Color = alternateColor; }
        gray = !gray;

        string info = $"Hour: {entry.Key * options.TimeMark}";
        hourRange.Value = info;
        index++;
    }

    hourRange = rawDataSheet.Range["F1"];
    hourRange.EntireColumn.ColumnWidth = 20;
}
private void FormatColumns(_Worksheet rawDataSheet)
{
    // Autofit columns
    Range[] ranges = new Range[4] { rawDataSheet.Range["A1"], rawDataSheet.Range["B1"], rawDataSheet.Range["C1"], rawDataSheet.Range["D1"] };
    foreach (Range r in ranges)
        r.EntireColumn.AutoFit();

    // Set formats
    Range formatRange = rawDataSheet.Range["A:B"];
    formatRange.NumberFormat = "[h]:mm:ss";
    formatRange = rawDataSheet.Range["C:C"];
    formatRange.NumberFormat = "General";
}
private void WriteStatsTable(_Worksheet sheet, string name, Stats stats, Dictionary<int, string> statesMapping, int position, bool isTotal = false)
{
    WriteHeader(sheet, name, position, isTotal);
    WritePhases(sheet, statesMapping, stats, position + 1, isTotal);
    WriteStats(sheet, stats, statesMapping, position + 1, isTotal);
}
private void WriteHeader(_Worksheet sheet, string title, int position, bool isTotal = false)
{
    sheet.Cells[position, 1] = title;
    sheet.Cells[position, 2] = "sec";
    sheet.Cells[position, 3] = "min";
    sheet.Cells[position, 4] = "%";
    sheet.Cells[position, 5] = "num";

    Range startCell = sheet.Cells[position, 1];
    Range endCell = sheet.Cells[position, 5];
    sheet.Range[startCell, endCell].Interior.Color = isTotal? secondaryDark : secondaryLight;
    startCell = sheet.Cells[position, 2];
    sheet.Range[startCell, endCell].HorizontalAlignment = XlHAlign.xlHAlignRight;
}
private void WritePhases(_Worksheet sheet, Dictionary<int, string> phases, Stats stats, int position, bool isTotal = false)
{
    int currPosition = position;
    for (int i = phases.Count; i > 0; i--)
    {
        sheet.Cells[currPosition, 1] = phases[i];
        currPosition++;
    }

    if (isTotal) { sheet.Cells[currPosition, 1] = "Total"; currPosition++; }

    // Color regualr phases
    int criteriaStartPos = currPosition;
    Range startCell = sheet.Cells[position, 1];
    Range endCell = sheet.Cells[currPosition, 1];
    sheet.Range[startCell, endCell].Interior.Color = isTotal ? primaryDark : primaryLight;

    // Write specific criteria phases (example "Wakfeulness < 5")
    foreach (KeyValuePair<int, int> entry in options.StateAndCriteria)
    {
        sheet.Cells[currPosition, 1] = $"{phases[entry.Key]} <= {entry.Value}s";
        additionalDistanceFromSpecificCrietria++;
        currPosition++;
    }
    // Write specific criteria phases (example "Wakfeulness > 5")
    foreach (KeyValuePair<int, int> entry in options.StateAndCriteriaAbove)
    {
        sheet.Cells[currPosition, 1] = $"{phases[entry.Key]} >= {entry.Value}s";
        additionalDistanceFromSpecificCrietria++;
        currPosition++;
    }

    //if (!isTotal) { currPosition--; }
    currPosition--;

    // Color specific criteria phases
    startCell = sheet.Cells[criteriaStartPos, 1];
    endCell = sheet.Cells[currPosition, 1];
    sheet.Range[startCell, endCell].Interior.Color = isTotal ? criteriaDark : criteriaLight;
}
private void WriteStats(_Worksheet sheet, Stats stats, Dictionary<int, string> phases, int position, bool isTotal = false)
{
    int curPosition = position;
    for (int i = phases.Count; i > 0; i--)
    {
        sheet.Cells[curPosition, 2] = stats.StateTimes[i];
        sheet.Cells[curPosition, 3] = Math.Round((double)stats.StateTimes[i] / 60, 2);
        sheet.Cells[curPosition, 4] = stats.SatePercentages[i];
        sheet.Cells[curPosition, 5] = stats.StateNumber[i];
        curPosition++;
    }

    // For total table
    if (isTotal) 
    {
        sheet.Cells[curPosition, 2] = stats.TotalTime;
        sheet.Cells[curPosition, 3] = Math.Round((double)stats.TotalTime / 60, 2);
        curPosition++;
    }

    // Specific criteria values
    foreach (KeyValuePair<int, int> CriteriaStateAndBelow in options.StateAndCriteria)
    {
        sheet.Cells[curPosition, 2] = stats.SpecificStateTimes[CriteriaStateAndBelow.Key];
        sheet.Cells[curPosition, 3] = Math.Round((double)stats.SpecificStateTimes[CriteriaStateAndBelow.Key] / 60, 2);
        sheet.Cells[curPosition, 5] = stats.SpecificTimeNumbers[CriteriaStateAndBelow.Key];
        curPosition++;
    }
    foreach (KeyValuePair<int, int> CriteriaStateAndAbove in options.StateAndCriteriaAbove)
    {
        sheet.Cells[curPosition, 2] = stats.SpecificStateTimesAbove[CriteriaStateAndAbove.Key];
        sheet.Cells[curPosition, 3] = Math.Round((double)stats.SpecificStateTimesAbove[CriteriaStateAndAbove.Key] / 60, 2);
        sheet.Cells[curPosition, 5] = stats.SpecificStateNumbersAbove[CriteriaStateAndAbove.Key];
        curPosition++;
    }
}
private void WriteGraphTable(_Worksheet sheet, string title, Dictionary<int, string> phases, int position, DataType type)
{
    WriteGraphHeader(sheet, title, phases, position);
    int columnPos = 2;
    int index = 1;
    int count = workfile.HourlyStats.Count;
    foreach (KeyValuePair<int, Stats> entry in workfile.HourlyStats)
    {
        updateStatus("Calculating graph stats", index, count);
        string header = entry.Value.TotalTime != options.TimeMark * 3600 && index == count ? $"Last {Math.Round(workfile.LastSectionTime / 60, 2)}hr - {workfile.LastSectionTime}min" : $"{entry.Key * options.TimeMark}hr";
        sheet.Cells[position, columnPos] = header;
        if (type == DataType.Percentage) { WriteGraphStats(sheet, entry.Value, phases, position + 1, columnPos, type); }
        else if (type == DataType.Minute) { WriteGraphStats(sheet, entry.Value, phases, position + 1, columnPos, type); }
        else if (type == DataType.Second) { WriteGraphStats(sheet, entry.Value, phases, position + 1, columnPos, type); }
        else if (type == DataType.Number) { WriteGraphStats(sheet, entry.Value, phases, position + 1, columnPos, type); }
        columnPos++;
        index++;
    }
    ColorRange(sheet, sheet.Cells[position, 2], sheet.Cells[position, columnPos - 1], secondaryLight);
    Range range = sheet.Range[sheet.Cells[position, 2], sheet.Cells[position, columnPos - 1]];
    range.HorizontalAlignment = XlHAlign.xlHAlignRight;
    range = sheet.Cells[position, columnPos - 1];
    range.EntireColumn.AutoFit();
}
private void WriteGraphHeader(_Worksheet sheet, string title, Dictionary<int, string> phases, int position)
{
    sheet.Cells[position, 1] = title;
    int curPosition = position + 1;
    for (int i = phases.Count; i > 0; i--)
    {
        sheet.Cells[curPosition, 1] = phases[i];
        curPosition++;
    }

    ColorRange(sheet, sheet.Cells[position, 1], sheet.Cells[position, 1], secondaryLight);
    ColorRange(sheet, sheet.Cells[position + 1, 1], sheet.Cells[curPosition - 1, 1], primaryLight);
}
private void WriteGraphStats(_Worksheet sheet, Stats stats, Dictionary<int, string> phases, int positionRow, int positionColumn, DataType type)
{
    int curPositionRow = positionRow;
    for (int i = phases.Count; i > 0; i--)
    {
        if (type == DataType.Percentage) { sheet.Cells[curPositionRow, positionColumn] = stats.SatePercentages[i]; }
        else if (type == DataType.Minute) { sheet.Cells[curPositionRow, positionColumn] = Math.Round((double)stats.StateTimes[i] / 60, 2); }
        else if (type == DataType.Second) { sheet.Cells[curPositionRow, positionColumn] = stats.StateTimes[i]; }
        else if (type == DataType.Number) { sheet.Cells[curPositionRow, positionColumn] = stats.StateNumber[i]; }
        curPositionRow++;
    }
}
private void CreateColumnChart(_Worksheet sheet, Stats stats)
{

}
private void WriteDuplicate(_Worksheet sheet, Tuple<int, int> duplicate, int position)
{
    sheet.Cells[position, 1] = duplicate.Item1;
    sheet.Cells[position, 2] = duplicate.Item2;
}
*/
    }
}
