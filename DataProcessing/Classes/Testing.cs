using DataProcessing.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    class Testing
    {
        public async Task ExportToExcelBackup(List<DataSample> records)
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
                    DataSample record = records[i];
                    sheet.Cells[index, 1] = record.AT.ToString();
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
