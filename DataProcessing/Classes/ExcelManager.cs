using DataProcessing.Classes.Calculate;
using DataProcessing.Classes.Export;
using DataProcessing.Models;
using DataProcessing.Utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    internal class ExcelManager
    {
        #region Private attributes
        // Vertical distance between tables in each collection
        private const int DISTANCE_BETWEEN_TABLES = 2;

        // Method to kill excel process (Otherwise excel process gets hung up in background)
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        // User selected export options
        private readonly CalculationOptions options;
        private readonly TableCreator _tableCreator;

        // Global services
        private readonly Services services = Services.GetInstance();

        // Reusable excel component references
        private _Workbook wb;
        private _Worksheet sheet;
        private int sheetNumber;
        private Range formatRange;
        #endregion

        #region Constructor
        // Constructor for importing
        public ExcelManager()
        {

        }
        // Constructor for exporting
        public ExcelManager(CalculationOptions calcOptions, CalculatedData data)
        {
            this.options = calcOptions;
            this._tableCreator = new TableCreator(calcOptions, data);
        }
        #endregion

        #region Public methods
        public async Task<Dictionary<int, ExcelSheetErrors>> CheckExcelFile(string filePath)
        {
            Dictionary<int, ExcelSheetErrors> errorsInSheets = new Dictionary<int, ExcelSheetErrors>();

            services.SetWorkStatus(true);

            await Task.Run(async () =>
            {
                // 1. Open excel
                services.UpdateWorkStatus("Opening excel...");
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(filePath, ReadOnly: false);
                Worksheet worksheet = workbook.Sheets[1];

                try
                {
                    int currentSheetNumber = 1;
                    ExcelSheetErrors errors;
                    foreach (Worksheet curSheet in workbook.Sheets)
                    {
                        errors = await CheckExcelSheet(curSheet);
                        if (errors.Count() > 0)
                        {
                            errorsInSheets.Add(currentSheetNumber, errors);
                        }
                        currentSheetNumber++;
                    }
                }
                catch(Exception e)
                {
                    if (e.Message == "First row can not be blank in excel file!")
                    {
                        Marshal.ReleaseComObject(worksheet);
                        Marshal.ReleaseComObject(workbook);
                        excel.Workbooks.Close();
                        excel.Quit();
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GetExcelProcess(excel).Kill();
                        services.SetWorkStatus(false);
                        throw e;
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

            return errorsInSheets;
        }
        public async Task HighlightExcelFileErrors(string filePath, Dictionary<int, ExcelSheetErrors> errorsInSheet)
        {
            await Task.Run(() =>
            {
                // 1. Open excel
                services.UpdateWorkStatus("Opening excel...");
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(filePath, ReadOnly: false);
                string errorsLogPath = Path.Combine(Path.GetDirectoryName(filePath), "errors.txt");
                File.WriteAllText(errorsLogPath, String.Empty);

                Worksheet worksheet;
                Range errorRange;
                Interior interior;
                // 1. Highlight main data errors
                foreach (KeyValuePair<int, ExcelSheetErrors> sheetAndErrors in errorsInSheet)
                {
                    worksheet = workbook.Sheets[sheetAndErrors.Key];
                    // 1. Highlight main data errors
                    foreach (int errorRow in sheetAndErrors.Value.MainDataErrorRows)
                    {
                        errorRange = worksheet.Range[$"A{errorRow}:B{errorRow}"];
                        interior = errorRange.Interior;
                        interior.Color = ExcelResources.GetInstance().Colors["DarkRed"];
                    }

                    // 2. Highlight behavior data errors
                    foreach (KeyValuePair<int, List<int>> entry in sheetAndErrors.Value.BehaviorErrorRows)
                    {
                        foreach (int errorRow in entry.Value)
                        {
                            errorRange = worksheet.Cells[errorRow + 1, entry.Key + 1];
                            interior = errorRange.Interior;
                            interior.Color = ExcelResources.GetInstance().Colors["DarkRed"];
                        }
                    }

                    
                    using (StreamWriter sw = File.AppendText(errorsLogPath))
                    {
                        sw.WriteLine(sheetAndErrors.Value.ErrorLog);
                    }

                }

                Process.Start("notepad.exe", errorsLogPath);

                excel.Visible = true;
                excel.UserControl = true;
                services.SetWorkStatus(false);
            });
        }
        public async Task ImportFromExcel(string filePath)
        {
            await Task.Run(() =>
            {
                // 1. Open excel
                services.UpdateWorkStatus("Starting import...");
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(filePath, ReadOnly: false);
                Worksheet worksheet;

                List<TimeStamp> dataToImport;
                int currentSheetNumber = 1;
                try
                {                    
                    for (int i = 1; i <= workbook.Sheets.Count; i++)
                    {
                        worksheet = workbook.Sheets[i];
                        dataToImport = GetImportableDataFromExcelSheet(worksheet);
                        Behaviours behaviours = GetBehaviorsFromExcelSheet(worksheet);
                        // If sheet also contains behaviour data append it to data
                        if (behaviours.Count() > 0)
                        {
                            dataToImport = AppendBehaviorsToData(dataToImport, behaviours);
                        }

                        // Persist to database
                        Services.GetInstance().UpdateWorkStatus($"Persisting data");
                        TimeStamp.SaveMany(dataToImport, currentSheetNumber);
                        currentSheetNumber++;
                    }
                }
                catch(Exception e)
                {
                    Marshal.ReleaseComObject(workbook);
                    excel.Workbooks.Close();
                    excel.Quit();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GetExcelProcess(excel).Kill();
                    services.SetWorkStatus(false);
                    throw e;
                }

                // Supposed to clean excel from memory but fails miserably
                Process excelProcess = GetExcelProcess(excel);
                excel.Workbooks.Close();
                while (Marshal.ReleaseComObject(workbook) > 0) ;
                workbook = null;
                while (Marshal.ReleaseComObject(excel.Workbooks) > 0) ;
                excel.Quit();
                while (Marshal.ReleaseComObject(excel) > 0) ;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                excelProcess.Kill();
            });
        }
        public async Task ExportToExcelC()
        {
            services.SetWorkStatus(true);
            await Task.Run(() =>
            {
                // 1. Open excel
                Application excel = new Application();
                excel.Caption = WorkfileManager.GetInstance().SelectedWorkFile.Name;
                wb = excel.Workbooks.Add(Missing.Value);

                // Create mandatory sheets
                CreateRawDataSheet();
                CreateStatsSheet();
                CreateGraphSheet();
                CreateDuplicatesSheet();
                CreateFrequenciesSheet();
                // Frequency ranges and cluster are optional so depedning on whether user selects them or not sheent position might differ
                if (options.FrequencyRanges.Count > 0) { CreateCustomFrequenciesSheet(); }
                if (options.ClusterSeparationTimeInSeconds > 0) { CreateClusterSheet(); }
                CreateBehaviorSheet();

                // Release excel to accesible state for user
                wb.Sheets[1].Select(Type.Missing);
                excel.Visible = true;
                excel.UserControl = true;
            });
            services.SetWorkStatus(false);
        }
        public async Task<int> CountSheets(string filePath)
        {
            int result = 0;

            await Task.Run(() =>
            {
                Application excel = new Application();
                Workbook workbook = excel.Workbooks.Open(filePath, ReadOnly: true);
                result = workbook.Sheets.Count;
                GetExcelProcess(excel).Kill();
            });

            return result;
        }
        #endregion

        #region Private helpers
        private async Task<ExcelSheetErrors> CheckExcelSheet(Worksheet sheet)
        {
            ExcelSheetErrors result = new ExcelSheetErrors();

            await Task.Run(() =>
            {
                Range firstRow = sheet.Cells[1, 1];
                if (firstRow.Value == null)
                {
                    throw new Exception("First row can not be blank in excel file!");
                }

                // 2. Count nonblank rows
                services.UpdateWorkStatus("Counting rows...");
                Range targetCells = sheet.UsedRange;
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
                Range rangeData = sheet.Range[$"A1:B{rowCount}"];
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
                // If parsing data from importing file failed (for example corrupted time entry) then we want to throw error on which row the parse failed
                catch (Exception e)
                {
                    Exception myException = new Exception(e.Message);
                    myException.Data.Add("Iteration", iteration);
                    throw myException;
                }

                // Check time integrity
                for (int i = 1; i < timeData.Count - 2; i++)
                {
                    if (!isBetweenTimeInterval(timeData[i - 1].Item1, timeData[i + 1].Item1, timeData[i].Item1))
                    {
                        result.AddMainDataErrorRow(i + 1);
                    }
                }

                // Check behavior integrity
                // We need two checks for behaviors, first in each interval the first time span has to be lower than the second. The other check is that the end of each behavior needs to be the start of another behavior or sleep.

                // 1. Check if there is any behaviors
                Behaviours behaviours = GetBehaviorsFromExcelSheet(sheet);

                if (behaviours.Count() == 0)
                {
                    return;
                }

                List<TimeSpan> sleepTimes = timeData.Where(ts => ts.Item2 == 2).Select(ts => ts.Item1).ToList();
                string errorLog;
                result.BehaviorErrorRows = behaviours.GetErrorRowIndexes(sleepTimes, out errorLog);
                result.ErrorLog = sheet.Name + "\n" + errorLog;
            });

            return result;
        }
        private List<TimeStamp> GetImportableDataFromExcelSheet(Worksheet sheet)
        {
            if (sheet.Cells[1, 1].Value == null)
            {
                throw new Exception("First row can not be blank in excel file!");
            }

            // 2. Count nonblank rows
            services.UpdateWorkStatus("Counting rows...");
            Range targetRange = sheet.UsedRange;
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
            targetRange = sheet.Range[$"A1:B{rowCount}"];
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

            // Supposed to clean excel from memory but fails miserably
            while (Marshal.ReleaseComObject(allRows) > 0) ;
            while (Marshal.ReleaseComObject(targetCells) > 0) ;
            while (Marshal.ReleaseComObject(targetRange) > 0) ;
            allRows = null;
            targetCells = null;
            targetRange = null;
            sheet = null;

            return markAddedSamples;
        }
        private Behaviours GetBehaviorsFromExcelSheet(Worksheet sheet)
        {
            services.UpdateWorkStatus("Importing behaviours...");
            Behaviours behaviours = new Behaviours();
            int rowNumber;
            string curRowData;
            string[] times;
            int[] timeUnits;
            TimeInterval curInterval;
            object curRowRawData;
            for (int i = 1; i <= 5; i++)
            {
                rowNumber = 2;
                while (true) 
                {
                    curRowRawData = sheet.Cells[rowNumber, i + 3].Value;
                    curRowData = Convert.ToString(curRowRawData);
                    if (String.IsNullOrEmpty(curRowData))
                    {
                        break;
                    }
                    times = curRowData.Split('-');
                    timeUnits = times[0].Split(':').Select(Int32.Parse).ToArray();
                    curInterval = new TimeInterval();
                    curInterval.From = new TimeSpan(timeUnits[0], timeUnits[1], timeUnits[2]);
                    timeUnits = times[1].Split(':').Select(Int32.Parse).ToArray();
                    curInterval.Till = new TimeSpan(timeUnits[0], timeUnits[1], timeUnits[2]);
                    behaviours.AddTimeIntervalOfBehaviour(curInterval, i + 2);
                    rowNumber++;
                }                        
            }

            return behaviours;
        }
        private List<TimeStamp> AppendBehaviorsToData(List<TimeStamp> data, Behaviours behaviours) 
        {
            List<TimeStamp> result = new List<TimeStamp>();

            TimeStamp prevStamp = data[0];
            result.Add(prevStamp);
            TimeStamp curStamp;
            int state;
            TimeInterval curInterval;
            List<Tuple<int, TimeInterval>> intervalsWithinWakefulness;
            TimeStamp behaviourStamp;
            for (int i = 1; i < data.Count; i++)
            {
                curStamp = data[i];
                state = curStamp.State;
                // Might not want to hard code it here!!!!!!
                if (state != 2)
                {
                    result.Add(curStamp);
                    prevStamp = curStamp;
                    continue;
                }
                curInterval = new TimeInterval();
                curInterval.From = prevStamp.Time;
                curInterval.Till = curStamp.Time;
                intervalsWithinWakefulness = behaviours.GetIntervalsWithinInvterval(curInterval);
                foreach (Tuple<int, TimeInterval> behaviourInterval in intervalsWithinWakefulness)
                {
                    behaviourStamp = new TimeStamp() { Time = behaviourInterval.Item2.Till, State = behaviourInterval.Item1 };
                    result.Add(behaviourStamp);
                }
                prevStamp = curStamp;
            }

            return result;
        }
        // Methods for creating each sheet (may include additional formatting and chart creating)
        private void CreateRawDataSheet()
        {
            _Worksheet sheet = wb.ActiveSheet;
            sheet.Name = "Raw Data";

            // Raw data
            services.UpdateWorkStatus("Exporting raw data");
            _tableCreator.CreateRawDataTable().ExportToSheet(sheet, 1, 1);

            formatRange = sheet.Range["A:B"];
            formatRange.NumberFormat = "[h]:mm:ss";
            formatRange = sheet.Range["C:C"];
            formatRange.NumberFormat = "General";

            // Latency table
            services.UpdateWorkStatus("Exporting latency");
            _tableCreator.CreateLatencyTable().ExportToSheet(sheet, 1, 8);

            formatRange = sheet.Range["H1:H1"];
            formatRange.EntireColumn.AutoFit();

            sheetNumber = 1;
        }
        private void CreateStatsSheet()
        {
            // Stat tables
            services.UpdateWorkStatus("Exporting stat tables");
            sheet = CreateNewSheet(wb, "Stats", sheetNumber);
            int vPos = 1;
            // Since all stat table has a chart beside it we have to add additional
            // distance between them, because the chart height is longer than the table
            // but if table has specific criterias (which will make the table longer)
            // than the additional increment will be smaller or sometimes nonexistent
            int criteriaNumber = options.Criterias.Count(c => c.Value != null);
            int additionaDistance = criteriaNumber >= 3 ? 0 : 3 - criteriaNumber;

            foreach (IExportable table in _tableCreator.CreateStatTables())
            {
                vPos = table.ExportToSheet(sheet, vPos, 1);
                vPos += DISTANCE_BETWEEN_TABLES + additionaDistance;
            }

            // Autofit first column
            formatRange = sheet.Range["A1:A1"];
            formatRange.EntireColumn.AutoFit();

            sheetNumber++;
        }
        private void CreateGraphSheet()
        {
            // Graph tables
            services.UpdateWorkStatus("Exporting graph tables");
            sheet = CreateNewSheet(wb, "Graph Stats", sheetNumber);
            int vPos = 1;
            foreach (IExportable table in _tableCreator.CreateGraphTables())
            {
                vPos = table.ExportToSheet(sheet, vPos, 1);
                vPos += DISTANCE_BETWEEN_TABLES;
            }

            // Autofit first column
            formatRange = sheet.Range["A1:A1"];
            formatRange.EntireColumn.AutoFit();

            sheetNumber++;
        }
        private void CreateDuplicatesSheet()
        {
            // Duplicates
            services.UpdateWorkStatus("Exporting duplicates");
            sheet = CreateNewSheet(wb, "Duplicated Graph Stats", sheetNumber);
            _tableCreator.CreateDuplicatesTable().ExportToSheet(sheet, 1, 1);

            sheetNumber++;
        }
        private void CreateFrequenciesSheet()
        {
            // Frequencies
            services.UpdateWorkStatus("Exporting frequencies");
            sheet = CreateNewSheet(wb, "Frequencies", sheetNumber);
            int vPos = 1;
            foreach (IExportable table in _tableCreator.CreateFrequencyTables())
            {
                vPos = table.ExportToSheet(sheet, vPos, 1);
                vPos += DISTANCE_BETWEEN_TABLES;
            }

            sheetNumber++;
        }
        private void CreateCustomFrequenciesSheet()
        {
            services.UpdateWorkStatus("Exporting custom frequency ranges");
            sheet = CreateNewSheet(wb, "Frequency ranges", sheetNumber);
            formatRange = sheet.Range["A:A"];
            formatRange.NumberFormat = "@";
            int vPos = 1;
            foreach (IExportable table in _tableCreator.CreateCustomFrequencyTables())
            {
                vPos = table.ExportToSheet(sheet, vPos, 1);
                vPos += DISTANCE_BETWEEN_TABLES;
            }

            sheetNumber++;
        }
        private void CreateClusterSheet()
        {
            services.UpdateWorkStatus("Exporting cluster data");
            sheet = CreateNewSheet(wb, "Clusters", sheetNumber);
            _tableCreator.CreateClusterDataTable().ExportToSheet(sheet, 1, 1);
            int vPos = 1;
            foreach (IExportable table in _tableCreator.CreateGraphTablesForClusters())
            {
                vPos = table.ExportToSheet(sheet, vPos, 5);
                vPos += DISTANCE_BETWEEN_TABLES;
            }

            formatRange = sheet.Range["E1:E1"];
            formatRange.EntireColumn.AutoFit();

            sheetNumber++;
        }
        private void CreateBehaviorSheet()
        {
            // Stat tables
            services.UpdateWorkStatus("Exporting stat tables");
            sheet = CreateNewSheet(wb, "Behavior", sheetNumber);
            int vPos = 1;
            int additionaDistance = 7;

            foreach (IExportable table in _tableCreator.CreateBehaviorStatTables())
            {
                vPos = table.ExportToSheet(sheet, vPos, 1);
                vPos += DISTANCE_BETWEEN_TABLES + additionaDistance;
            }

            // Autofit first column
            formatRange = sheet.Range["A1:A1"];
            formatRange.EntireColumn.AutoFit();

            sheetNumber++;
        }

        // Creates new sheet in excel file
        private _Worksheet CreateNewSheet(_Workbook workbook, string name, int position)
        {
            Sheets sheets = workbook.Sheets;
            Worksheet sheet = sheets.Add(Type.Missing, sheets[position], Type.Missing, Type.Missing);
            sheet.Name = name;
            return sheet;
        }
        // Used in import, we update working status only once every 10 iteration so it won't be overloaded
        private void updateStatus(string label, int index, int count)
        {
            if (index % 10 == 0 || index == count) { services.UpdateWorkStatus($"{label} {index}/{count}"); }
        }
        // Checks if timespan is in given interval
        private bool isBetweenTimeInterval(TimeSpan from, TimeSpan till, TimeSpan time)
        {
            return from <= time && time <= till;
            if (from < till)
            {
                return from <= time && time <= till;
            }
            else
            {
                return from <= time || time <= till;
            }
        }
        // Shortcut to get range in excel sheet
        private Range GetRange(_Worksheet sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Range start = sheet.Cells[startRow, startColumn];
            Range end = sheet.Cells[endRow, endColumn];
            return sheet.Range[start, end];
        }
        // Gets excel process to later kill (otherwise it gets hung up as a background process)
        private Process GetExcelProcess(Application excelApp)
        {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }
        #endregion
    }
}
