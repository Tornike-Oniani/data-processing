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
    internal class ExcelManager
    {
        private const int DISTANCE_BETWEEN_TABLES = 2;

        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        private Dictionary<string, Color> colors = new Dictionary<string, Color>()
        {
            {"DarkBlue", Color.FromArgb(75, 177, 250) },
            {"DarkOrange", Color.FromArgb(250, 148, 75) },
            {"DarkRed", Color.FromArgb(250, 92, 75) },
            {"Blue", Color.FromArgb(148, 216, 255) },
            {"Orange", Color.FromArgb(255, 187, 148) },
            {"Green", Color.FromArgb(202, 255, 138) },
            {"Yellow", Color.FromArgb(250, 228, 102) },
            {"Gray", Color.FromArgb(230, 229, 225) },
            {"Red", Color.FromArgb(255, 157, 148) }
        };

        private TableCollection RawData;
        private TableCollection Latency;
        private TableCollection Stats;
        private TableCollection Graphs;
        private TableCollection Duplicates;
        private TableCollection Frequencies;
        private TableCollection FrequencyRanges;
        private TableCollection ClusterData;
        private TableCollection ClusterGraphs;

        private ExportOptions options;
        private Workfile workfile = WorkfileManager.GetInstance().SelectedWorkFile;
        private Services services = Services.GetInstance();

        public ExcelManager(
            ExportOptions options,
            TableCollection RawData,
            TableCollection Latency,
            TableCollection Stats,
            TableCollection Graphs,
            TableCollection Duplicates,
            TableCollection Frequencies,
            TableCollection FrequencyRanges,
            TableCollection ClusterData,
            TableCollection ClusterGraphs
            )
        {
            this.options = options;
            this.RawData = RawData;
            this.Latency = Latency;
            this.Stats = Stats;
            this.Graphs = Graphs;
            this.Duplicates = Duplicates;
            this.Frequencies = Frequencies;
            this.FrequencyRanges = FrequencyRanges;
            this.ClusterData = ClusterData;
            this.ClusterGraphs = ClusterGraphs;
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
                int iteration = 0;
                try
                {
                    for (int i = 1; i <= dataValues.Length / 2; i++)
                    {
                        iteration = i;
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
                }
                // If parsing data from importing file failed (for example corrupted time entry) then we want to throw arrow on which row the parse failed
                catch (Exception e)
                {
                    Exception myException = new Exception(e.Message);
                    myException.Data.Add("Iteration", iteration);
                    GetExcelProcess(excel).Kill();
                    services.SetWorkStatus(false);
                    throw myException;
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
                // Kill process from task manager
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
                    interior.Color = colors["DarkRed"];
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

        public async Task ExportToExcelC()
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

                // Raw data
                services.UpdateWorkStatus("Exporting raw data");
                ExportTableCollection(rawDataSheet, RawData, 1);

                Range formatRange = rawDataSheet.Range["A:B"];
                formatRange.NumberFormat = "[h]:mm:ss";
                formatRange = rawDataSheet.Range["C:C"];
                formatRange.NumberFormat = "General";

                // Latency table
                services.UpdateWorkStatus("Exporting latency");
                ExportTableCollection(rawDataSheet, Latency, 8);

                // Stat tables
                _Worksheet statSheet = CreateNewSheet(wb, "Stats", 1);
                services.UpdateWorkStatus("Exporting stat tables");
                ExportTableCollection(statSheet, Stats, 1);

                wb.Sheets[1].Select(Type.Missing);
                excel.Visible = true;
                excel.UserControl = true;
            });
            services.SetWorkStatus(false);
        }

        // Export whole table collection on excel sheet
        private void ExportTableCollection(_Worksheet sheet, TableCollection collection, int horizontalPosition)
        {
            object[,] tableArray;
            int rowPos = 1;
            Range range;
            // Keep track of number of iteration in foreach loop
            int counter = 0;

            // Write each table in collection into sheet
            foreach (System.Data.DataTable table in collection.Tables)
            {
                // Convert table to 2d array
                tableArray = DataTableTo2DArray(table, collection.HasHeader, collection.HasTiteOnTop);
                // Get appropriate range in excel and set its value to 2d array
                range = GetRange(sheet, rowPos, horizontalPosition, tableArray.GetLength(0), horizontalPosition - 1 + tableArray.GetLength(1));
                range.Value = tableArray;

                // Decorate table
                decorateTable(sheet, collection.ColorRanges, horizontalPosition, rowPos, counter == 0 && collection.HasTotal);

                // Prepare row position for the next table
                rowPos += tableArray.GetLength(0) + DISTANCE_BETWEEN_TABLES;
            }
        }

        // We want to convert data table to object array because excel works fastest when you select a range and set it's value to 2d array
        private object[,] DataTableTo2DArray(System.Data.DataTable table, bool includeColumnNames, bool titleOnTop)
        {
            int index = 0;
            int rowNumber = table.Rows.Count;

            if (includeColumnNames) { rowNumber++; }
            if (titleOnTop) { rowNumber++; }

            // Rows + 1 because we also have to add column names and title
            object[,] table2D = new object[rowNumber, table.Columns.Count];

            // Title
            table2D[0, 0] = table.TableName;

            // If we want title on top we have to incement index so all other values come below it
            if (titleOnTop) { index++; }

            // Column names
            if (includeColumnNames)
            {
                for (int i = index; i < table.Columns.Count; i++)
                {
                    table2D[index, i] = table.Columns[i].ColumnName;
                }
                index++;
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
                    table2D[i + index, j] = curRow[curColumn.ColumnName];
                }
            }

            return table2D;
        }
        private void decorateTable(_Worksheet sheet, Dictionary<string, ColorRange[]> colorRanges, int horizontalPosition, int verticalPosition, bool useDarkColors)
        {
            string colorName;
            ColorRange[] ranges;
            Color color;
            Range excRange;
            int startRow;
            int startColumn;
            int endRow;
            int endColumn;

            foreach (KeyValuePair<string, ColorRange[]> entry in colorRanges)
            {
                colorName = entry.Key;
                ranges = entry.Value;
                // Get appropriate color from dictionary (if we are decorating total table we will use dark version of the color
                color = colors[(useDarkColors ? "Dark" : "") + colorName];

                foreach (ColorRange range in ranges)
                {
                    // Set relative positions (ColorRange keeps track of ranges relative to table disregarding current position on excel)
                    startRow = range.StartRow + verticalPosition;
                    startColumn = range.StartColumn + horizontalPosition;
                    endRow = range.EndRow + verticalPosition;
                    endColumn = range.EndColumn + horizontalPosition;

                    // Get range and set its color
                    excRange = GetRange(sheet, startRow, startColumn, endRow, endColumn);
                    excRange.Interior.Color = color;
                }
            }
        }

        private _Worksheet CreateNewSheet(_Workbook workbook, string name, int position)
        {
            Sheets sheets = workbook.Sheets;
            Worksheet sheet = sheets.Add(Type.Missing, sheets[position], Type.Missing, Type.Missing);
            sheet.Name = name;
            return sheet;
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
        private Range GetRange(_Worksheet sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Range start = sheet.Cells[startRow, startColumn];
            Range end = sheet.Cells[endRow, endColumn];
            return sheet.Range[start, end];
        }
        private Process GetExcelProcess(Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }


        /*
        public async Task ExportToExcel()
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

                int prev = 1;
                formatRange = rawDataSheet.Range[$"A{prev}:E{prev}"];
                formatRange.Interior.Color = timeMarkColor;
                Tuple<int, string> indexTime;
                for (int i = 0; i < hourRowIndexesTime.Count; i++)
                {
                    indexTime = hourRowIndexesTime[i];
                    formatRange = rawDataSheet.Range[$"A{indexTime.Item1}:E{indexTime.Item1}"];
                    formatRange.Interior.Color = timeMarkColor;
                    formatRange = rawDataSheet.Range[$"F{prev}:F{indexTime.Item1}"];
                    formatRange.Merge();
                    formatRange.Value = indexTime.Item2;
                    formatRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    formatRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    if (i % 2 == 0) formatRange.Interior.Color = alternateColor;
                    prev = indexTime.Item1 + 1;

                    // Autofit after finishing
                    if (i == hourRowIndexesTime.Count - 1)
                        formatRange.EntireColumn.AutoFit();
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

                // Write latency table
                WriteLatencyTable(rawDataSheet, latencyTable, 8);

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

                // Frequency table
                services.UpdateWorkStatus("Exporting frequencies");
                _Worksheet frequencySheet = CreateNewSheet(wb, "Frequenies", 4);
                frequencySheet.EnableSelection = XlEnableSelection.xlNoSelection;
                int pos = 1;
                foreach (DataTableInfo tableInfo in frequenciesCollection)
                {
                    pos = WriteFrequencyTable(frequencySheet, tableInfo, pos);
                }

                // Autofit first column (In case the total hour is not round we will get 'Last x minutes' in the last table and we need autofit for that)
                Range start = frequencySheet.Cells[1, 1];
                start.EntireColumn.AutoFit();

                // Autofit every second column (ones which have 'frequency' in title)
                for (int i = 1; i <= frequenciesCollection[0].Table.Columns.Count; i++)
                {
                    if (i % 2 == 0)
                    {
                        start = frequencySheet.Cells[1, i];
                        start.EntireColumn.AutoFit();
                    }
                }

                // Custom frequency ranges
                if (customFrequenciesCollection != null)
                {
                    services.UpdateWorkStatus("Exporting custom frequency ranges");
                    _Worksheet customFrequencySheet = CreateNewSheet(wb, "Frequency ranges", 5);
                    frequencySheet.EnableSelection = XlEnableSelection.xlNoSelection;
                    pos = 1;
                    foreach (DataTableInfo tableInfo in customFrequenciesCollection)
                    {
                        pos = WriteCustomFrequencyTable(customFrequencySheet, tableInfo, pos);
                        if (tableInfo.IsTotal)
                        {
                            // chartTop = WriteStatChart(statsSheet, tableInfo, 5, 10, tableCount, chartTop);
                            ChartObjects charts = customFrequencySheet.ChartObjects();
                            ChartObject chartObject = charts.Add(chartLeft, 1, 17 * cellWidth, 22 * cellHeight);
                            Chart chart = chartObject.Chart;

                            //range = GetRange(customFrequencySheet, 3, 1, tableInfo.Table.Rows.Count + 2, tableInfo.Table.Columns.Count - 2);
                            range = GetRange(customFrequencySheet, 2, 1, tableInfo.Table.Rows.Count + 2, tableInfo.Table.Columns.Count - 2);
                            chart.ChartWizard(
                                range,
                                XlChartType.xlColumnClustered,
                                Title: tableInfo.Table.TableName,
                                ValueTitle: "Frequency");
                            foreach (Series series in chart.SeriesCollection())
                            {
                                series.HasDataLabels = true;
                            }
                            chart.HasLegend = true;
                            chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;
                            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                            //range = GetRange(customFrequencySheet, 2, 1, 2, tableInfo.Table.Columns.Count - 2);
                            //xAxis.CategoryNames = range;
                        }
                    }

                    // Autofit first column (In case the total hour is not round we will get 'Last x minutes' in the last table and we need autofit for that)
                    start = customFrequencySheet.Cells[1, 1];
                    start.EntireColumn.AutoFit();

                    // Autofit every second column (ones which have 'frequency' in title)
                    for (int i = 1; i <= customFrequenciesCollection[0].Table.Columns.Count; i++)
                    {
                        if (i % 2 == 0)
                        {
                            start = customFrequencySheet.Cells[1, i];
                            start.EntireColumn.AutoFit();
                        }
                    }
                }

                // If cluster separation time was set other than 0 we want to generate cluster data
                if (options.ClusterSeparationTimeInSeconds != 0) 
                {
                    // Clusters
                    services.UpdateWorkStatus("Exporting clusters");
                    _Worksheet clusterSheet = CreateNewSheet(wb, "Clusters", 5);
                    frequencySheet.EnableSelection = XlEnableSelection.xlNoSelection;

                    // Write cluster raw data
                    WriteClusterData(clusterSheet, nonMarkedTimeStamps, 1);

                    // Mark cluster ends with red color
                    for (int i = 0; i < clusterLocation.Count; i++)
                    {
                        formatRange = clusterSheet.Range[$"A{clusterLocation[i]}:B{clusterLocation[i]}"];
                        formatRange.Interior.Color = clusterColor;
                    }

                    // Color each phase in cluster with appropriate color
                    for (int i = 0; i < clusterColorIndexState.Count; i++)
                    {
                        formatRange = clusterSheet.Range[$"A{clusterColorIndexState[i].Item1}:B{clusterColorIndexState[i].Item1}"];
                        // We have array that corresponds each state to color, but since indexes in array start from 0 we have to substract 1 from state for mapping
                        formatRange.Interior.Color = clusterColorsForEachState[clusterColorIndexState[i].Item2 - 1];
                    }

                    // Create graph charts
                    position = 1;
                    index = 1;
                    max = graphTableCollectionForClusters.Count;
                    foreach (DataTableInfo tableInfo in graphTableCollectionForClusters)
                    {
                        updateStatus("Graph tables for clusters", index, max);
                        position = WriteDataTable(clusterSheet, tableInfo, position, 5);
                        lastHour = graphSheet.Cells[position, tableInfo.HeaderIndexes.Item2];
                        lastHour.EntireColumn.AutoFit();
                        position += distanceBetweenTables;
                        index++;
                    }

                    clusterSheet.Range["E1"].EntireColumn.AutoFit();

                    range = graphSheet.Range["A1:K1"];
                    graphLeft = range.Width;
                    // Width is calculated by columns for now (parameter 10 doesn't do anything)
                    WriteGraphChartForCluster(clusterSheet, graphTableCollectionForClusters[0], 15, 1);
                }

                wb.Sheets[1].Select(Type.Missing);
                excel.Visible = true;
                excel.UserControl = true;
            });
            services.SetWorkStatus(false);
        }
         */
    }
}
