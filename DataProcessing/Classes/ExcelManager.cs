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
        private Color grayLight = Color.FromArgb(242, 242, 242);
        private Color clusterColor = Color.FromArgb(222, 83, 49);
        // Colors for state coloring in each cluster Indexes: 0 - PS, 1 - sleep, 2 - wakefulness 
        private Color[] clusterColorsForEachState = new Color[3] 
        {
            Color.FromArgb(188, 255, 130), 
            Color.FromArgb(255, 236, 115),             
            Color.FromArgb(255, 182, 163)};
        private List<Tuple<int, int>> clusterColorIndexState = new List<Tuple<int, int>>();
        private Workfile workfile = WorkfileManager.GetInstance().SelectedWorkFile;
        private int distanceBetweenTables = 2;
        private ExportOptions options;
        private List<DataTableInfo> statTableCollection;
        private List<DataTableInfo> graphTableCollection;
        private List<DataTableInfo> graphTableCollectionForClusters;
        private List<DataTableInfo> frequenciesCollection;
        private DataTableInfo latencyTable;
        List<DataTableInfo> customFrequenciesCollection;
        private Services services = Services.GetInstance();
        private List<int> markerLocations = new List<int>();
        private List<int> darkLightMarkerLocations = new List<int>();
        private List<int> clusterLocation = new List<int>();
        private double cellWidth;
        private double cellHeight;
        private double chartLeft;
        private double chartLeftAlt;
        private double chartVerticalDistance;
        private double graphLeft;
        private double duplicatedChartLeft;
        List<int> statTablePositions = new List<int>();

        public ExcelManager(
            ExportOptions options, 
            List<DataTableInfo> statTableCollection, 
            List<DataTableInfo> graphTableCollection,
            List<DataTableInfo> graphTableCollectionForClusters,
            List<DataTableInfo> frequenciesCollection,
            DataTableInfo latencyTable,
            List<DataTableInfo> customFrequenciesCollection
            )
        {
            this.options = options;
            this.statTableCollection = statTableCollection;
            this.graphTableCollection = graphTableCollection;
            this.graphTableCollectionForClusters = graphTableCollectionForClusters;
            this.frequenciesCollection = frequenciesCollection;
            this.latencyTable = latencyTable;
            this.customFrequenciesCollection = customFrequenciesCollection;
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
        public async Task ExportToExcel(
            List<TimeStamp> timeStamps,
            List<TimeStamp> nonMarkedTimeStamps,
            List<Tuple<int, int>> duplicatedTimes, 
            List<int> hourRowIndexes,
            List<Tuple<int, string>> hourRowIndexesTime)
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

        private int WriteDataTable(_Worksheet sheet, DataTableInfo tableInfo, int position, int horizontalPosition = 1)
        {
            int curPosition = position;
            System.Data.DataTable table = tableInfo.Table;

            object[,] values = DataTableTo2DArray(table);
            Range start = sheet.Cells[curPosition, horizontalPosition];
            Range end = sheet.Cells[curPosition + table.Rows.Count, horizontalPosition + table.Columns.Count - 1];
            Range tableRange = sheet.Range[start, end];
            tableRange.Value = values;

            // Color header
            start = sheet.Cells[position, horizontalPosition];
            end = sheet.Cells[position, horizontalPosition + tableInfo.HeaderIndexes.Item2 - 1];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? secondaryDark : secondaryLight;

            // Color phases
            start = sheet.Cells[position + tableInfo.PhasesIndexes.Item1, horizontalPosition];
            end = sheet.Cells[position + tableInfo.PhasesIndexes.Item2, horizontalPosition];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? primaryDark : primaryLight;

            // Color criteria
            if (tableInfo.CriteriaPhases != null)
            {
                if (tableInfo.CriteriaPhases.Item2 > 0)
                {
                    start = sheet.Cells[position + tableInfo.CriteriaPhases.Item1 + 1, horizontalPosition];
                    end = sheet.Cells[position + tableInfo.CriteriaPhases.Item1 + tableInfo.CriteriaPhases.Item2, horizontalPosition];
                    tableRange = sheet.Range[start, end];
                    tableRange.Interior.Color = tableInfo.IsTotal ? criteriaDark : criteriaLight;
                }
            }

            // Set alignments
            start = sheet.Cells[position, horizontalPosition + 1];
            end = sheet.Cells[position, horizontalPosition + tableInfo.HeaderIndexes.Item2 - 1];
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
        private void WriteClusterData(_Worksheet sheet, List<TimeStamp> timeStamps, int position)
        {
            object[,] values = ListTo2DArrayNonFull(timeStamps);
            Range start = sheet.Cells[position, 1];
            Range end = sheet.Cells[position + timeStamps.Count - 1, 2];
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
            // Get range to determine width of chart
            Range range = GetRange(sheet, 1, 1, 1, tableInfo.Table.Columns.Count);
            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(0, graphTableCollection.Count * 6 * cellHeight, range.Width, height * cellHeight);
            Chart chart = chartObject.Chart;

            range = GetRange(sheet, tablePos + 1, 1, tablePos + options.MaxStates, tableInfo.Table.Columns.Count);
            chart.ChartWizard(
                range,
                XlChartType.xlColumnClustered,
                PlotBy: XlRowCol.xlRows,
                Title: tableInfo.Table.TableName,
                ValueTitle: "Percents");
            chart.HasLegend = true;
            chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            range = GetRange(sheet, tablePos, 2, tablePos, tableInfo.Table.Columns.Count);
            xAxis.CategoryNames = range;
        }
        private void WriteGraphChartForCluster(_Worksheet sheet, DataTableInfo tableInfo, int height, int tablePos)
        {
            // Get range to determine width of chart
            Range range = GetRange(sheet, 1, 5, 1, 5 + tableInfo.Table.Columns.Count);
            Range leftPos = GetRange(sheet, 1, 1, 1, 4);
            ChartObjects charts = sheet.ChartObjects();
            ChartObject chartObject = charts.Add(leftPos.Width, graphTableCollectionForClusters.Count * 6 * cellHeight, range.Width, height * cellHeight);
            Chart chart = chartObject.Chart;

            // Get range of data table to map it to chart
            range = GetRange(sheet, tablePos + 1, 5, tablePos + options.MaxStates, 5 + tableInfo.Table.Columns.Count);
            chart.ChartWizard(
                range,
                XlChartType.xlColumnClustered,
                PlotBy: XlRowCol.xlRows,
                Title: tableInfo.Table.TableName,
                ValueTitle: "Percents");
            chart.HasLegend = true;
            chart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;
            Axis xAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
            range = GetRange(sheet, tablePos, 6, tablePos, 5 + tableInfo.Table.Columns.Count);
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
        private int WriteFrequencyTable(_Worksheet sheet, DataTableInfo tableInfo, int position)
        {
            object[,] values = FrequencyDataTableTo2DArray(tableInfo.Table);
            Range start = sheet.Cells[position, 1];
            Range end = sheet.Cells[position + tableInfo.Table.Rows.Count + 1, tableInfo.Table.Columns.Count];
            Range tableRange = sheet.Range[start, end];
            tableRange.Value = values;

            // Color header
            start = sheet.Cells[position, 1];
            end = sheet.Cells[position, 1];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? secondaryDark : secondaryLight;

            // Color phases
            start = sheet.Cells[position + 1, 1];
            end = sheet.Cells[position + 1, tableInfo.Table.Columns.Count];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? primaryDark : primaryLight;

            return position + tableInfo.Table.Rows.Count + 3;
        }
        private int WriteCustomFrequencyTable(_Worksheet sheet, DataTableInfo tableInfo, int position)
        {
            object[,] values = CustomFrequencyDataTableTo2DArray(tableInfo.Table);
            Range start = sheet.Cells[position, 1];
            Range end = sheet.Cells[position + tableInfo.Table.Rows.Count + 1, tableInfo.Table.Columns.Count - 2];
            Range tableRange = sheet.Range[start, end];
            tableRange.NumberFormat = "@";
            tableRange.Value = values;

            // Color header
            start = sheet.Cells[position, 1];
            end = sheet.Cells[position, 1];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? secondaryDark : secondaryLight;

            // Color phases
            start = sheet.Cells[position + 1, 1];
            end = sheet.Cells[position + 1, tableInfo.Table.Columns.Count - 2];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? primaryDark : primaryLight;

            // Color times
            //for (int i = 1; i <= tableInfo.Table.Columns.Count; i++)
            //{
            //    if (i % 2 != 0)
            //    {
            //        start = sheet.Cells[position + 2, i];
            //        end = sheet.Cells[position + 1 + tableInfo.Table.Rows.Count, i];
            //        tableRange = sheet.Range[start, end];
            //        tableRange.Interior.Color = grayLight;
            //    }
            //}
            start = sheet.Cells[position + 2, 1];
            end = sheet.Cells[position + 1 + tableInfo.Table.Rows.Count, 1];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = grayLight;

            return position + tableInfo.Table.Rows.Count + 3;
        }
        private void WriteLatencyTable(_Worksheet sheet, DataTableInfo tableInfo, int position)
        {
            object[,] values = FrequencyDataTableTo2DArray(tableInfo.Table);
            Range start = sheet.Cells[1, position];
            Range end = sheet.Cells[tableInfo.Table.Rows.Count + 2, position + tableInfo.Table.Columns.Count - 1];
            Range tableRange = sheet.Range[start, end];
            tableRange.Value = values;

            // Color header
            start = sheet.Cells[1, position];
            end = sheet.Cells[1, position];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? secondaryDark : secondaryLight;

            // Color phases
            start = sheet.Cells[2, position];
            end = sheet.Cells[2, position + tableInfo.Table.Columns.Count - 1];
            tableRange = sheet.Range[start, end];
            tableRange.Interior.Color = tableInfo.IsTotal ? primaryDark : primaryLight;

            // Autofit
            start.EntireColumn.AutoFit();
            end.EntireColumn.AutoFit();
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
        private object[,] FrequencyDataTableTo2DArray(System.Data.DataTable table)
        {
            // Rows + 1 because we also have to add column names and title
            object[,] table2D = new object[table.Rows.Count + 2, table.Columns.Count];

            // Title
            table2D[0, 0] = table.TableName;

            // Column names
            for (int i = 0; i < table.Columns.Count; i++)
            {
                table2D[1, i] = table.Columns[i].ColumnName;
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
                    if ((int)curRow[curColumn.ColumnName] == 0)
                        table2D[i + 2, j] = null;
                    else
                        table2D[i + 2, j] = curRow[curColumn.ColumnName];
                }
            }

            return table2D;
        }
        private object[,] CustomFrequencyDataTableTo2DArray(System.Data.DataTable table)
        {
            // Rows + 1 because we also have to add column names and title
            object[,] table2D = new object[table.Rows.Count + 2, table.Columns.Count];

            // Title
            table2D[0, 0] = table.TableName;

            int tempIndex = 0;
            // Column names
            for (int i = 0; i < table.Columns.Count; i++)
            {
                // First column are frequency ranges
                if (i == 0)
                {
                    table2D[1, tempIndex] = "Ranges";
                    tempIndex++;
                }
                if (i % 2 != 0)
                {
                    table2D[1, tempIndex] = table.Columns[i].ColumnName;
                    tempIndex++;
                }
            }

            // Table contents
            DataRow curRow;
            DataColumn curColumn;
            for (int i = 0; i < table.Rows.Count; i++)
            {
                tempIndex = 0;
                curRow = table.Rows[i];
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    if (j == 0 || j % 2 != 0)
                    {
                        curColumn = table.Columns[j];
                        table2D[i + 2, tempIndex] = curRow[curColumn.ColumnName];
                        tempIndex++;
                    }
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
        // We use this for cluster data, because in cluster data all we need is seconds and states not the other properties of timestamp
        private object[,] ListTo2DArrayNonFull(List<TimeStamp> timeStamps)
        {
            object[,] list2D = new object[timeStamps.Count, 5];

            TimeStamp curTimeStamp;
            for (int i = 0; i < timeStamps.Count; i++)
            {
                curTimeStamp = timeStamps[i];
                // If we encounter wakefulness that is more than our cluster specified time then mark its index for coloring
                if (curTimeStamp.TimeDifferenceInSeconds > options.ClusterSeparationTimeInSeconds && curTimeStamp.State == 3) 
                { clusterLocation.Add(i + 1); }
                // Otherwise add to regular state-based coloring list for cluster
                else if (curTimeStamp.State != 0) 
                { clusterColorIndexState.Add(new Tuple<int, int>(i + 1, curTimeStamp.State)); }
                list2D[i, 0] = curTimeStamp.TimeDifferenceInSeconds;
                list2D[i, 1] = curTimeStamp.State;
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
        private Process GetExcelProcess(Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }
    }
}
