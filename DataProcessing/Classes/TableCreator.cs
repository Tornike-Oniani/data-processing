using DataProcessing.Classes.Export;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace DataProcessing.Classes
{
    /// <summary>
    /// Takes processed data and creates all tables that are to be exported
    /// </summary>
    internal class TableCreator
    {
        // Data that was calculated by DataProcessor
        private CalculatedData calculatedData;
        // List of timestamps for raw data
        private List<TimeStamp> timeStamps;
        // List of timestamps for cluster data
        private List<TimeStamp> nonMarkedTimeStamps;
        // Exported options selected by user
        private ExportOptions options;
        // Table decorator for coloring
        private TableDecorator decorator;

        // Constructor
        public TableCreator(CalculatedData data, ExportOptions options, List<TimeStamp> timeStamps, List<TimeStamp> nonMarkedTimeStamps)
        {
            // Init
            this.calculatedData = data;
            this.options = options;
            this.timeStamps = timeStamps;
            this.nonMarkedTimeStamps = nonMarkedTimeStamps;
            decorator = new TableDecorator();
        }

        // Public table collection creators
        public ExcelTable CreateRawDataTable()
        {
            int rowCount = timeStamps.Count;
            int colCount = 5;
            object[,] table = new object[rowCount, colCount];

            // Fill in data
            int rowIndex = 0;
            foreach (TimeStamp timeStamp in timeStamps)
            {
                // We convert time into string other wise setting range value in excel won't work
                table[rowIndex, 0] = timeStamp.Time.ToString();
                table[rowIndex, 1] = timeStamp.TimeDifference.ToString();
                table[rowIndex, 2] = timeStamp.TimeDifferenceInDouble;
                table[rowIndex, 3] = timeStamp.TimeDifferenceInSeconds;
                table[rowIndex, 4] = timeStamp.State;
                rowIndex++;
            }

            // Decorate collection and return it
            return decorator.DecorateRawData(new ExcelTable(table), timeStamps);
        }
        public ExcelTable CreateLatencyTable()
        {
            List<DataTable> tables = new List<DataTable>();
            DataTable table = new DataTable("Latency");

            // Add columns
            table.Columns.Add(new DataColumn("First sleep", typeof(int)));
            table.Columns.Add(new DataColumn("First PS", typeof(int)));

            // Fill in data
            DataRow row = table.NewRow();
            row["First sleep"] = calculatedData.timeBeforeFirstSleep;
            row["First PS"] = calculatedData.timeBeforeFirstParadoxicalSleep;
            table.Rows.Add(row);

            tables.Add(table);

            // Decorate table and return it
            return decorator.DecorateLatencyTable(tables);
        }
        public TableCollection CreateStatTables()
        {
            List<DataTable> tables = new List<DataTable>();

            // Create a total stat table and add it on top
            tables.Add(CreateStatTable("Total", calculatedData.totalStats, true));

            // Add table for each hour mark
            int counter = 1;
            string tableName = "";
            foreach (KeyValuePair<int, Stats> hourAndStat in calculatedData.hourAndStats)
            {
                tableName = $"hour {hourAndStat.Key * options.TimeMark}";

                // If we reached the stats for final hour and it is not full hour then we name the table with remaning time
                if (counter == calculatedData.hourAndStats.Count && hourAndStat.Value.TotalTime % 3600 != 0)
                {
                    tableName = getTimeForStats(hourAndStat.Value.TotalTime);
                }

                tables.Add(CreateStatTable(tableName, hourAndStat.Value, false));
                counter++;
            }

            // Decorate collection and return it
            int criteriaNumber = options.Criterias.Where(c => c.Value != null).Count();
            return decorator.DecorateStatTables(tables, criteriaNumber);
        }
        public TableCollection CreateGraphTables()
        {
            List<DataTable> tables = new List<DataTable>();

            // Create graph tables
            tables.Add(CreateGraphTable("Percentages %", GraphTableDataType.Percentages));
            tables.Add(CreateGraphTable("Minutes", GraphTableDataType.Minutes));
            tables.Add(CreateGraphTable("Seconds", GraphTableDataType.Seconds));
            tables.Add(CreateGraphTable("Numbers", GraphTableDataType.Numbers));

            // Decorate collection and return it
            return decorator.DecorateGraphTables(tables);
        }
        public TableCollection CreateDuplicatesTable()
        {
            List<DataTable> tables = new List<DataTable>();
            DataTable table = new DataTable();
            DataRow row;

            // Add columns
            table.Columns.Add(new DataColumn("Time", typeof(int)));
            table.Columns.Add(new DataColumn("State", typeof(int)));

            // Fill in data
            foreach (Tuple<int, int> item in calculatedData.duplicatedTimes)
            {
                row = table.NewRow();
                row["Time"] = item.Item1;
                row["State"] = item.Item2;
                table.Rows.Add(row);
            }

            tables.Add(table);

            // Decorate collection and return it
            return decorator.DecorateDuplicatesTable(tables);
        }
        public TableCollection CreateFrequencyTables()
        {
            List<DataTable> tables = new List<DataTable>();

            int hour = 0;
            foreach (Dictionary<int, SortedList<int, int>> stateTimeFrequency in calculatedData.hourStateFrequencies)
            {
                // First item will be total not hourly
                if (hour == 0)
                {
                    tables.Add(CreateFrequencyTable("Total", stateTimeFrequency, true));
                    hour++;
                    continue;
                }

                tables.Add(CreateFrequencyTable($"hour {hour}", stateTimeFrequency));
                hour++;
            }

            return decorator.DecorateFrequencyTables(tables);
        }
        public TableCollection CreateCustomFrequencyTables()
        {
            // If user didn't provide any custom frequency
            if (options.customFrequencyRanges.Count == 0) { return null; }

            List<DataTable> tables = new List<DataTable>();

            int hour = 0;
            foreach (Dictionary<int, Dictionary<string, int>> stateTimeFrequency in calculatedData.hourStateCustomFrequencies)
            {
                // First item will be total not hourly
                if (hour == 0)
                {
                    tables.Add(CreateCustomFrequencyTable("Total", stateTimeFrequency, true));
                    hour++;
                    continue;
                }

                tables.Add(CreateCustomFrequencyTable($"hour {hour}", stateTimeFrequency));
                hour++;
            }

            return decorator.DecorateCustomFrequencyTables(tables, options.customFrequencyRanges.Count);
        }
        public TableCollection CreateClusterDataTable()
        {
            List<DataTable> tables = new List<DataTable>();
            DataTable table = new DataTable();
            DataRow row = table.NewRow();

            // Add columns
            table.Columns.Add(new DataColumn("Time", typeof(int)));
            table.Columns.Add(new DataColumn("State", typeof(int)));

            // Fill in data
            foreach (TimeStamp timeStamp in nonMarkedTimeStamps)
            {
                row["Time"] = timeStamp.TimeDifferenceInSeconds;
                row["State"] = timeStamp.State;
                table.Rows.Add(row);
                row = table.NewRow();
            }

            tables.Add(table);

            // Decorate collection and return it
            return decorator.DecorateClusterDataTable(tables, nonMarkedTimeStamps, options.ClusterSeparationTimeInSeconds);
        }
        public TableCollection CreateGraphTablesForClusters()
        {
            List<DataTable> tables = new List<DataTable>();

            // Create graph tables
            tables.Add(CreateGraphTableForCluster("Percentages %", GraphTableDataType.Percentages));
            tables.Add(CreateGraphTableForCluster("Minutes", GraphTableDataType.Minutes));
            tables.Add(CreateGraphTableForCluster("Seconds", GraphTableDataType.Seconds));
            tables.Add(CreateGraphTableForCluster("Numbers", GraphTableDataType.Numbers));

            // Decorate collection and return it
            return decorator.DecorateGraphTables(tables);
        }

        // Single table helper creators
        private object[,] CreateStatTable(string name, Stats stats, bool isTotal)
        {
            // Header + all the phases + optional criterias
            int rowCount =
                calculatedData.stateAndPhases.Count +
                options.Criterias.Count(c => c.Value != null) +
                1;
            object[,] table = new object[rowCount, 5];

            // Set title
            table[0, 0] = name;

            // Add columns
            table[0, 1] = "sec";
            table[0, 2] = "min";
            table[0, 3] = "%";
            table[0, 4] = "num";

            // Fill in essential data
            // We start from 1 because 0 is set to header
            int rowIndex = 1;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                table[rowIndex, 0] = stateAndPhase.Value;
                table[rowIndex, 1] = stats.StateTimes[stateAndPhase.Key];
                table[rowIndex, 2] = Math.Round((double)stats.StateTimes[stateAndPhase.Key] / 60, 2);
                table[rowIndex, 3] = stats.StatePercentages[stateAndPhase.Key];
                table[rowIndex, 4] = stats.StateNumber[stateAndPhase.Key];
                rowIndex++;
            }

            // If table is total add one more row for summed up stats
            if (isTotal)
            {
                table[rowIndex, 0] = "Total";
                table[rowIndex, 1] = stats.TotalTime;
                table[rowIndex, 2] = Math.Round((double)stats.TotalTime / 60, 2);
                rowIndex++;
            }

            // Fill in data for specific criterias if any was set
            foreach (SpecificCriteria criteria in options.Criterias)
            {
                // Skip nonexistent criterias
                if (criteria.Value == null) { continue; }

                table[rowIndex, 0] = $"{calculatedData.stateAndPhases[criteria.State]} {criteria.GetOperandValue()} {criteria.Value}";
                table[rowIndex, 1] = stats.SpecificTimes[criteria];
                table[rowIndex, 2] = Math.Round((double)stats.SpecificTimes[criteria] / 60, 2);
                table[rowIndex, 3] = stats.SpecificNumbers[criteria];
                rowIndex++;
            }

            return table;
        }
        // Division can be either hourAndStats or clusterAndStats
        private object[,] CreateGraphTable(string name, Dictionary<int, Stats> division, GraphTableDataType dataType)
        {
            // Header + phases
            int rowCount = calculatedData.stateAndPhases.Count + 1;
            // Phases + each hour mark
            int colCount = calculatedData.hourAndStats.Count + 1;
            object[,] table = new object[rowCount, colCount];

            // Set title
            table[0, 0] = name;

            // Add columns
            // We start from 1 because 0 is set to title
            int colIndex = 1;
            foreach (KeyValuePair<int, Stats> hourAndStat in division)
            {
                table[0, colIndex] = $"{colIndex}ep";
                colIndex++;
            }

            // Fill in data
            // We start from 1 because 0 is set to header
            int rowIndex = 1;
            colIndex = 0;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                table[rowIndex, colIndex] = stateAndPhase.Value;
                colIndex++;
                switch (dataType)
                {
                    case GraphTableDataType.Seconds:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            table[rowIndex, colIndex] = hourAndStat.Value.StateTimes[stateAndPhase.Key];
                            colIndex++;
                        }
                        break;
                    case GraphTableDataType.Minutes:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            table[rowIndex, colIndex] = Math.Round((double)hourAndStat.Value.StateTimes[stateAndPhase.Key] / 60, 2);
                            colIndex++;
                        }
                        break;
                    case GraphTableDataType.Percentages:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            table[rowIndex, colIndex] = hourAndStat.Value.StatePercentages[stateAndPhase.Key];
                            colIndex++;
                        }
                        break;
                    case GraphTableDataType.Numbers:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            table[rowIndex, colIndex] = hourAndStat.Value.StateNumber[stateAndPhase.Key];
                            colIndex++;
                        }
                        break;
                }
                colIndex = 0;
                rowIndex++;
            }

            return table;
        }
        private object[,] CreateFrequencyTable(string name, Dictionary<int, SortedList<int, int>> stateFrequencies, bool isTotal = false)
        {
            // Find largest dictionary to iterate with
            int max = 0;
            foreach (KeyValuePair<int, SortedList<int, int>> stateTimeFrequency in stateFrequencies)
            {
                if (stateTimeFrequency.Value.Count > max) { max = stateTimeFrequency.Value.Count; }
            }

            // largest dictionary + title + header
            int rowCount = max + 2;
            // time and frequency for each state
            int colCount = options.MaxStates * 2;
            object[,] table = new object[rowCount, colCount];

            // Set title
            table[0, 0] = name;

            // Add columns based on states
            int colIndex = 0;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                table[1, colIndex] = $"{stateAndPhase.Value.Substring(0, 1)} time";
                table[1, colIndex + 1] = $"{stateAndPhase.Value.Substring(0, 1)} freq";
                colIndex += 2;
            }

            // Add data from dictionary to table

            int? time;
            int? frequency;
            SortedList<int, int> current;
            // We start from 2 because 0 is set to title and 1 is set to header
            int rowIndex = 2;
            colIndex = 0;
            // Iterate with largest dictionary (let's say Wakefulness has more variety than others, its dictionary will be larger)
            for (int i = 0; i < max; i++)
            {
                foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
                {
                    current = stateFrequencies[stateAndPhase.Key];
                    // Since we are going with largest dictionary there will be cases that index is out of range for smaller ones
                    // DEPRECATED: in that case we just set time and frequency to 0, which we will convert into blank/null during excel export
                    // Since we divided responsibilites we have to assign null values right here
                    if (i < current.Count)
                    {
                        time = current.ElementAt(i).Key;
                        frequency = current.ElementAt(i).Value;
                        table[rowIndex, colIndex] = time;
                        table[rowIndex, colIndex + 1] = frequency;
                    }
                    else
                    {
                        // Since we can't pass null to datatable column we have to use DBNull.Value instead
                        table[rowIndex, colIndex] = null;
                        table[rowIndex, colIndex + 1] = null;
                    }
                    colIndex += 2;
                }
                rowIndex++;
                colIndex = 0;
            }

            return table;
        }
        private object[,] CreateCustomFrequencyTable(string name, Dictionary<int, Dictionary<string, int>> stateCustomFrequencies, bool isTotal = false)
        {
            int max = 0;
            foreach (KeyValuePair<int, Dictionary<string, int>> stateTimeFrequency in stateCustomFrequencies)
            {
                if (stateTimeFrequency.Value.Count > max) { max = stateTimeFrequency.Value.Count; }
            }

            // Frequenc ranges + title + header
            int rowCount = max + 2;
            // Ranges column + frequency for each state
            int colCount = options.MaxStates + 1;
            object[,] table = new object[rowCount, colCount];

            // Set title
            table[0, 0] = name;

            // Add columns based on states
            table[1, 0] = "Ranges";
            int colIndex = 1;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                table[1, colIndex] = $"{stateAndPhase.Value.Substring(0, 1)} freq";
                colIndex++;
            }

            string range;
            int frequency;
            Dictionary<string, int> current;
            int rowIndex = 2;
            colIndex = 0;
            for (int i = 0; i < max; i++)
            {
                foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
                {
                    current = stateCustomFrequencies[stateAndPhase.Key];
                    range = current.ElementAt(i).Key;
                    frequency = current.ElementAt(i).Value;
                    table[rowIndex, 0] = range;
                    table[rowIndex, colIndex] = frequency;
                    colIndex++;
                }
                rowIndex++;
            }

            return table;
        }

        // Small helper functions
        private string getCriteriaLabel(SpecificCriteria criteria)
        {
            return $"{calculatedData.stateAndPhases[criteria.State]} {criteria.GetOperandValue()} {criteria.Value}";
        }
        private string getTimeForStats(int seconds)
        {
            if (seconds % 3600 == 0) { return $"hour {seconds / 3600}"; }

            TimeSpan span = TimeSpan.FromSeconds(seconds);
            if (span.TotalHours < 1) { return $"Last {Math.Round(span.TotalMinutes)} minutes"; }
            return $"Last {span.Hours} hours and {span.Minutes} minutes";
        }
        private string getTimeForGraph(int seconds)
        {
            if (seconds % 3600 == 0) { return $"{seconds / 3600}hr"; }

            TimeSpan span = TimeSpan.FromSeconds(seconds);
            if (span.TotalHours < 1) { return $"Last {Math.Round(span.TotalMinutes)} minutes"; }
            return $"Last {span.Hours}hr {span.Minutes} min";
        }
    }
}
