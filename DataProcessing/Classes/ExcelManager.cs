using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        private int additionalDistanceFromSpecificCrietria = 0;
        private ExportOptions options;
        private Services services = Services.GetInstance();
        private List<int> markerLocations = new List<int>();

        public ExcelManager(ExportOptions options)
        {
            this.options = options;
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
                sheet.Cells[curPosition, 4] = stats.TimePercentages[i];
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
                if (type == DataType.Percentage) { sheet.Cells[curPositionRow, positionColumn] = stats.TimePercentages[i]; }
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
        Process GetExcelProcess(Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        public async Task ExportToExcelBackup(List<TimeStamp> records)
        {
            Services services = Services.GetInstance();
            services.SetWorkStatus(true);

            await Task.Run(() =>
            {
                services.UpdateWorkStatus("Opening excel...");
                Application excel = new Application();
                _Workbook wb = excel.Workbooks.Add(Missing.Value);
                _Worksheet sheet = wb.ActiveSheet;
                sheet.Name = "Raw Data";
                excel.Caption = WorkfileManager.GetInstance().SelectedWorkFile.Name;

                int index = 1;
                int max = records.Count();
                for (int i = 0; i < max; i++)
                {
                    if (index % 10 == 0 || index == max) { services.UpdateWorkStatus($"Exporting data {index}/{max}"); }
                    TimeStamp record = records[i];
                    sheet.Cells[index, 1] = record.Time.ToString();
                    sheet.Cells[index, 5] = record.State;
                    index++;
                }

                Range[] ranges = new Range[4] { sheet.Range["A1"], sheet.Range["B1"], sheet.Range["C1"], sheet.Range["D1"] };
                foreach (Range range in ranges)
                {
                    range.EntireColumn.AutoFit();
                }

                Range formatRange = sheet.Range["A:A"];
                formatRange.NumberFormat = "[h]:mm:ss";
                formatRange = sheet.Range["B:B"];
                formatRange.NumberFormat = "[h]:mm:ss";

                // Set formulas
                services.UpdateWorkStatus("Setting formulas");
                int count = records.Count;
                sheet.Range["B2", $"B{count}"].Formula = "=IF(A2<A1, A2+1, A2)-A1";
                sheet.Range["C2", $"C{count}"].Formula = "=B2";
                sheet.Range["D2", $"D{count}"].Formula = "=C2*86400";

                formatRange = sheet.Range["C:C"];
                formatRange.NumberFormat = "General";

                // Setting calculations for Raw
                services.UpdateWorkStatus("Setting calculations");
                Dictionary<int, int[]> indexes = WorkfileManager.GetInstance().SelectedWorkFile.HourlyIndexes;
                Dictionary<int, Stats> dictStats = WorkfileManager.GetInstance().SelectedWorkFile.HourlyStats;
                Range statRange;
                index = 1;
                max = indexes.Count();
                foreach (KeyValuePair<int, int[]> entry in indexes)
                {
                    if (index % 10 == 0 || index == max) { services.UpdateWorkStatus($"Calculating {index}/{max}"); }
                    statRange = sheet.Range[$"F{entry.Value[0]}:F{entry.Value[1]}"];
                    statRange.Merge();
                    statRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    statRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                    if (entry.Key % 2 == 0)
                    {
                        sheet.Range[$"F{entry.Value[0]}:F{entry.Value[1]}"].Interior.Color = XlRgbColor.rgbLightGray;
                    }

                    Stats stats = dictStats[entry.Key];
                    string statInfo = $"Hour: {entry.Key}\n";
                    foreach (KeyValuePair<int, int> stateAndTime in stats.StateTimes)
                    {
                        statInfo += $"State {stateAndTime.Key}: {stateAndTime.Value} - {stats.TimePercentages[stateAndTime.Key]}%\n";
                    }
                    statRange.Value = statInfo;
                }

                Range fColumn = sheet.Range["F:F"];
                fColumn.EntireColumn.ColumnWidth = 20;

                // Set calculations
                services.UpdateWorkStatus("Calculating stats");
                Sheets sheets = wb.Sheets;
                Worksheet calcSheet = sheets.Add(Type.Missing, sheets[1], Type.Missing, Type.Missing);
                calcSheet.Name = "Stats";

                // Set total labels
                calcSheet.Cells[1, 1] = "Total";
                calcSheet.Cells[2, 1] = "Wakefulness";
                calcSheet.Cells[3, 1] = "Light sleep";
                calcSheet.Cells[4, 1] = "Deep sleep";
                calcSheet.Cells[5, 1] = "Paradoxical sleep";
                calcSheet.Cells[6, 1] = "Total time";
                calcSheet.Cells[1, 2] = "sec";
                calcSheet.Cells[1, 3] = "min";
                calcSheet.Cells[1, 4] = "%";

                calcSheet.Range[calcSheet.Cells[1, 2], calcSheet.Cells[1, 4]].HorizontalAlignment = XlHAlign.xlHAlignRight;

                // Set total stats
                Stats totlaStats = WorkfileManager.GetInstance().SelectedWorkFile.Stats;
                calcSheet.Cells[2, 2] = totlaStats.StateTimes[4];
                calcSheet.Cells[2, 3] = Math.Round((double)totlaStats.StateTimes[4] / 60, 2);
                calcSheet.Cells[2, 4] = totlaStats.TimePercentages[4];

                calcSheet.Cells[3, 2] = totlaStats.StateTimes[3];
                calcSheet.Cells[3, 3] = Math.Round((double)totlaStats.StateTimes[3] / 60, 2);
                calcSheet.Cells[3, 4] = totlaStats.TimePercentages[3];

                calcSheet.Cells[4, 2] = totlaStats.StateTimes[2];
                calcSheet.Cells[4, 3] = Math.Round((double)totlaStats.StateTimes[2] / 60, 2);
                calcSheet.Cells[4, 4] = totlaStats.TimePercentages[2];

                calcSheet.Cells[5, 2] = totlaStats.StateTimes[1];
                calcSheet.Cells[5, 3] = Math.Round((double)totlaStats.StateTimes[1] / 60, 2);
                calcSheet.Cells[5, 4] = totlaStats.TimePercentages[1];

                calcSheet.Cells[6, 2] = totlaStats.TotalTime;
                calcSheet.Cells[6, 3] = Math.Round((double)totlaStats.TotalTime / 60, 2);

                // Total stats coloring
                Range _coloring = calcSheet.Range["A1:D1"];
                _coloring.Interior.Color = System.Drawing.Color.FromArgb(250, 148, 75);
                _coloring = calcSheet.Range["A2:A6"];
                _coloring.Interior.Color = System.Drawing.Color.FromArgb(75, 177, 250);

                // Stats per hour
                int step = 8;
                index = 1;
                max = dictStats.Count();
                foreach (KeyValuePair<int, Stats> entry in dictStats)
                {
                    Range startCell;
                    Range endCell;
                    Range coloring;
                    if (index % 10 == 0 || index == max) { services.UpdateWorkStatus($"Calculating {index}/{max}"); }
                    // Set second, minute, percent columns
                    calcSheet.Cells[step, 2] = "sec";
                    calcSheet.Cells[step, 3] = "min";
                    calcSheet.Cells[step, 4] = "%";
                    calcSheet.Range[calcSheet.Cells[step, 2], calcSheet.Cells[step, 4]].HorizontalAlignment = XlHAlign.xlHAlignRight;

                    calcSheet.Cells[step, 1] = $"{entry.Key} hour";
                    coloring = calcSheet.Range[calcSheet.Cells[step, 1], calcSheet.Cells[step, 4]];
                    coloring.Interior.Color = System.Drawing.Color.FromArgb(255, 187, 148);
                    step++;
                    startCell = calcSheet.Cells[step, 1];
                    calcSheet.Cells[step, 1] = "Wakefulness";
                    calcSheet.Cells[step, 2] = entry.Value.StateTimes[4];
                    calcSheet.Cells[step, 3] = Math.Round((double)entry.Value.StateTimes[4] / 60, 2);
                    endCell = calcSheet.Cells[step, 4];
                    calcSheet.Cells[step, 4] = entry.Value.TimePercentages[4];
                    step++;
                    calcSheet.Cells[step, 1] = "Light sleep";
                    calcSheet.Cells[step, 2] = entry.Value.StateTimes[3];
                    calcSheet.Cells[step, 3] = Math.Round((double)entry.Value.StateTimes[3] / 60, 2);
                    calcSheet.Cells[step, 4] = entry.Value.TimePercentages[3];
                    step++;
                    calcSheet.Cells[step, 1] = "Deep sleep";
                    calcSheet.Cells[step, 2] = entry.Value.StateTimes[2];
                    calcSheet.Cells[step, 3] = Math.Round((double)entry.Value.StateTimes[2] / 60, 2);
                    calcSheet.Cells[step, 4] = entry.Value.TimePercentages[2];
                    step++;
                    endCell = calcSheet.Cells[step, 1];
                    calcSheet.Cells[step, 1] = "Paradoxical sleep";
                    calcSheet.Cells[step, 2] = entry.Value.StateTimes[1];
                    calcSheet.Cells[step, 3] = Math.Round((double)entry.Value.StateTimes[1] / 60, 2);
                    calcSheet.Cells[step, 4] = entry.Value.TimePercentages[1];
                    coloring = calcSheet.Range[startCell, endCell];
                    coloring.Interior.Color = System.Drawing.Color.FromArgb(148, 216, 255);
                    //coloring.Interior.Color = XlRgbColor.rgbLightBlue;
                    step += 2;
                    index++;
                }

                Range firstRange = calcSheet.Range[$"A1:A{max}"];
                firstRange.Columns.AutoFit();

                // Set graph calculations
                services.UpdateWorkStatus("Calculating graph stats");
                Worksheet calcGraphSheet = sheets.Add(Type.Missing, sheets[2], Type.Missing, Type.Missing);
                calcGraphSheet.Name = "Graph Stats";

                // Percentages graph
                // Set labels
                calcGraphSheet.Cells[1, 1] = "Percentages %";
                calcGraphSheet.Cells[2, 1] = "Wakefulness";
                calcGraphSheet.Cells[3, 1] = "Light sleep";
                calcGraphSheet.Cells[4, 1] = "Deep sleep";
                calcGraphSheet.Cells[5, 1] = "Paradoxical sleep";

                index = 1;
                max = dictStats.Count();
                step = 2;
                foreach (KeyValuePair<int, Stats> entry in dictStats)
                {
                    if (index % 10 == 0 || index == max) { services.UpdateWorkStatus($"Calculating {index}/{max}"); }
                    calcGraphSheet.Cells[1, step] = $"{entry.Key}hr";
                    calcGraphSheet.Cells[2, step] = entry.Value.TimePercentages[4];
                    calcGraphSheet.Cells[3, step] = entry.Value.TimePercentages[3];
                    calcGraphSheet.Cells[4, step] = entry.Value.TimePercentages[2];
                    calcGraphSheet.Cells[5, step] = entry.Value.TimePercentages[1];
                    step++;
                    index++;
                }

                firstRange = calcGraphSheet.Range[$"A1:A{max}"];
                firstRange.Columns.AutoFit();

                // Coloring
                Range header = calcGraphSheet.Range[calcGraphSheet.Cells[1, 2], calcGraphSheet.Cells[1, max + 1]];
                header.HorizontalAlignment = XlHAlign.xlHAlignRight;
                header.Interior.Color = System.Drawing.Color.FromArgb(148, 216, 255);
                header = calcGraphSheet.Range[$"A1:A{5}"];
                header.Interior.Color = System.Drawing.Color.FromArgb(148, 216, 255);
                header = calcGraphSheet.Range["A1:A1"];
                header.Interior.Color = System.Drawing.Color.FromArgb(255, 187, 148);

                // Minutes graph
                // Set labels
                int shift = 7;
                calcGraphSheet.Cells[shift + 1, 1] = "Minutes";
                calcGraphSheet.Cells[shift + 2, 1] = "Wakefulness";
                calcGraphSheet.Cells[shift + 3, 1] = "Light sleep";
                calcGraphSheet.Cells[shift + 4, 1] = "Deep sleep";
                calcGraphSheet.Cells[shift + 5, 1] = "Paradoxical sleep";

                index = 1;
                max = dictStats.Count();
                step = 2;
                foreach (KeyValuePair<int, Stats> entry in dictStats)
                {
                    if (index % 10 == 0 || index == max) { services.UpdateWorkStatus($"Calculating {index}/{max}"); }
                    calcGraphSheet.Cells[shift + 1, step] = $"{entry.Key}hr";
                    calcGraphSheet.Cells[shift + 2, step] = Math.Round((double)entry.Value.StateTimes[4] / 60, 2);
                    calcGraphSheet.Cells[shift + 3, step] = Math.Round((double)entry.Value.StateTimes[3] / 60, 2);
                    calcGraphSheet.Cells[shift + 4, step] = Math.Round((double)entry.Value.StateTimes[2] / 60, 2);
                    calcGraphSheet.Cells[shift + 5, step] = Math.Round((double)entry.Value.StateTimes[1] / 60, 2);
                    step++;
                    index++;
                }

                // Coloring
                header = calcGraphSheet.Range[calcGraphSheet.Cells[shift + 1, 2], calcGraphSheet.Cells[shift + 1, max + 1]];
                header.HorizontalAlignment = XlHAlign.xlHAlignRight;
                header.Interior.Color = System.Drawing.Color.FromArgb(148, 216, 255);
                header = calcGraphSheet.Range[$"A{shift + 1}:A{shift + 5}"];
                header.Interior.Color = System.Drawing.Color.FromArgb(148, 216, 255);
                header = calcGraphSheet.Range[$"A{shift + 1}:A{shift + 1}"];
                header.Interior.Color = System.Drawing.Color.FromArgb(255, 187, 148);

                // Seconds graph
                // Set labels
                shift = 14;
                calcGraphSheet.Cells[shift + 1, 1] = "Seconds";
                calcGraphSheet.Cells[shift + 2, 1] = "Wakefulness";
                calcGraphSheet.Cells[shift + 3, 1] = "Light sleep";
                calcGraphSheet.Cells[shift + 4, 1] = "Deep sleep";
                calcGraphSheet.Cells[shift + 5, 1] = "Paradoxical sleep";

                index = 1;
                max = dictStats.Count();
                step = 2;
                foreach (KeyValuePair<int, Stats> entry in dictStats)
                {
                    if (index % 10 == 0 || index == max) { services.UpdateWorkStatus($"Calculating {index}/{max}"); }
                    calcGraphSheet.Cells[shift + 1, step] = $"{entry.Key}hr";
                    calcGraphSheet.Cells[shift + 2, step] = entry.Value.StateTimes[4];
                    calcGraphSheet.Cells[shift + 3, step] = entry.Value.StateTimes[3];
                    calcGraphSheet.Cells[shift + 4, step] = entry.Value.StateTimes[2];
                    calcGraphSheet.Cells[shift + 5, step] = entry.Value.StateTimes[1];
                    step++;
                    index++;
                }

                // Coloring
                header = calcGraphSheet.Range[calcGraphSheet.Cells[shift + 1, 2], calcGraphSheet.Cells[shift + 1, max + 1]];
                header.HorizontalAlignment = XlHAlign.xlHAlignRight;
                header.Interior.Color = System.Drawing.Color.FromArgb(148, 216, 255);
                header = calcGraphSheet.Range[$"A{shift + 1}:A{shift + 5}"];
                header.Interior.Color = System.Drawing.Color.FromArgb(148, 216, 255);
                header = calcGraphSheet.Range[$"A{shift + 1}:A{shift + 1}"];
                header.Interior.Color = System.Drawing.Color.FromArgb(255, 187, 148);

                // Duplicated graph list
                services.UpdateWorkStatus("Duplicating graph stats");
                Worksheet calcDuplicateSheet = sheets.Add(Type.Missing, sheets[3], Type.Missing, Type.Missing);
                calcDuplicateSheet.Name = "Duplicated Graph Stats";

                step = 1;
                List<Tuple<int, int>> duplicatedStats = WorkfileManager.GetInstance().SelectedWorkFile.DuplicatedTimes;

                index = 1;
                max = duplicatedStats.Count();
                foreach (Tuple<int, int> duplicate in duplicatedStats)
                {
                    if (index % 10 == 0 || index == max) { services.UpdateWorkStatus($"Duplicating {index}/{max}"); }
                    calcDuplicateSheet.Cells[step, 1] = duplicate.Item1;
                    calcDuplicateSheet.Cells[step, 2] = duplicate.Item2;
                    step++;
                    index++;
                }

                // Finish
                sheet.Select(Type.Missing);
                excel.Visible = true;
                excel.UserControl = true;

            });

            services.SetWorkStatus(false);
        }
    }
}
