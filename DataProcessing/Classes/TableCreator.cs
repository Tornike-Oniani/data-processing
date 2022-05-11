using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public TableCreator(CalculatedData data, ExportOptions options, List<TimeStamp> timeStamps, List<TimeStamp> nonMarkedTimeStamps)
        {
            // Init
            this.calculatedData = data;
            this.options = options;
            this.timeStamps = timeStamps;
            this.nonMarkedTimeStamps = nonMarkedTimeStamps;
            decorator = new TableDecorator();
        }

        public TableCollection CreateRawDataTable()
        {
            List<DataTable> tables = new List<DataTable>();

            DataTable table = new DataTable("Raw data");
            DataRow row;

            // Add columns
            table.Columns.Add("Time", typeof(string));
            table.Columns.Add("Time difference", typeof(string));
            table.Columns.Add("Time difference in double", typeof(double));
            table.Columns.Add("Time difference in seconds", typeof(int));
            table.Columns.Add("State", typeof(int));

            // Fill in data
            foreach (TimeStamp timeStamp in timeStamps)
            {
                row = table.NewRow();
                // We convert time into string other wise setting range value in excel won't work
                row["Time"] = timeStamp.Time.ToString();
                row["Time difference"] = timeStamp.TimeDifference.ToString();
                row["Time difference in double"] = timeStamp.TimeDifferenceInDouble;
                row["Time difference in seconds"] = timeStamp.TimeDifferenceInSeconds;
                row["State"] = timeStamp.State;
                table.Rows.Add(row);
            }

            tables.Add(table);

            // Decorate collection and return it
            return decorator.DecorateRawData(tables, timeStamps);
        }
        public TableCollection CreateLatencyTable()
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
            return decorator.DecorateClusterDataTable(tables, timeStamps, options.ClusterSeparationTimeInSeconds);
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

        private DataTable CreateStatTable(string name, Stats stats, bool isTotal)
        {
            DataTable table = new DataTable(name);
            DataRow row;

            // Create columns
            table.Columns.Add(new DataColumn("Phases", typeof(string)));
            table.Columns.Add(new DataColumn("sec", typeof(int)));
            table.Columns.Add(new DataColumn("min", typeof(double)));
            table.Columns.Add(new DataColumn("%", typeof(double)));
            table.Columns.Add(new DataColumn("num", typeof(int)));

            // Fill in essential data
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                row = table.NewRow();
                row["Phases"] = stateAndPhase.Value;
                row["sec"] = stats.StateTimes[stateAndPhase.Key];
                row["min"] = Math.Round((double)stats.StateTimes[stateAndPhase.Key] / 60, 2);
                row["%"] = stats.StatePercentages[stateAndPhase.Key];
                row["num"] = stats.StateNumber[stateAndPhase.Key];
                table.Rows.Add(row);
            }

            // If table is total add one more row for summed up stats
            if (isTotal)
            {
                row = table.NewRow();
                row["Phases"] = "Total";
                row["sec"] = stats.TotalTime;
                row["min"] = Math.Round((double)stats.TotalTime / 60, 2);
                table.Rows.Add(row);
            }

            // Fill in data for specific criterias if any was set
            foreach (SpecificCriteria criteria in options.Criterias)
            {
                // Skip nonexistent criterias
                if (criteria.Value == null) { continue; }

                row = table.NewRow();
                row["Phases"] = $"{calculatedData.stateAndPhases[criteria.State]} {criteria.GetOperandValue()} {criteria.Value}";
                row["sec"] = stats.SpecificTimes[criteria];
                row["min"] = Math.Round((double)stats.SpecificTimes[criteria] / 60, 2);
                row["num"] = stats.SpecificNumbers[criteria];
                table.Rows.Add(row);
            }

            return table;
        }
        private DataTable CreateGraphTable(string name, GraphTableDataType dataType)
        {
            DataTable table = new DataTable(name);
            DataRow row;

            table.Columns.Add(new DataColumn("Phases", typeof(string)));
            Type columnType = typeof(string);
            switch (dataType)
            {
                case GraphTableDataType.Seconds: columnType = typeof(int); break;
                case GraphTableDataType.Minutes: columnType = typeof(double); break;
                case GraphTableDataType.Percentages: columnType = typeof(double); break;
                case GraphTableDataType.Numbers: columnType = typeof(int); break;
            }

            // Add columns
            int counter = 1;
            string tableName = "";
            foreach (KeyValuePair<int, Stats> hourAndStat in calculatedData.hourAndStats)
            {
                tableName = $"{hourAndStat.Key * options.TimeMark}hr";

                // If we reached the stats for final hour and it is not full hour then we name the table with remaning time
                if (counter == calculatedData.hourAndStats.Count && hourAndStat.Value.TotalTime % 3600 != 0)
                {
                    tableName = getTimeForGraph(hourAndStat.Value.TotalTime);
                }

                table.Columns.Add(new DataColumn(tableName, columnType));
                counter++;
            }

            // Fill in data
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                row = table.NewRow();
                row["Phases"] = stateAndPhase.Value;
                switch (dataType)
                {
                    case GraphTableDataType.Seconds:
                        foreach (KeyValuePair<int, Stats> hourAndStat in calculatedData.hourAndStats)
                        {
                            row[hourAndStat.Key] = hourAndStat.Value.StateTimes[stateAndPhase.Key];
                        }
                        break;
                    case GraphTableDataType.Minutes:
                        foreach (KeyValuePair<int, Stats> hourAndStat in calculatedData.hourAndStats)
                        {
                            row[hourAndStat.Key] = Math.Round((double)hourAndStat.Value.StateTimes[stateAndPhase.Key] / 60, 2);
                        }
                        break;
                    case GraphTableDataType.Percentages:
                        foreach (KeyValuePair<int, Stats> hourAndStat in calculatedData.hourAndStats)
                        {
                            row[hourAndStat.Key] = hourAndStat.Value.StatePercentages[stateAndPhase.Key];
                        }
                        break;
                    case GraphTableDataType.Numbers:
                        foreach (KeyValuePair<int, Stats> hourAndStat in calculatedData.hourAndStats)
                        {
                            row[hourAndStat.Key] = hourAndStat.Value.StateNumber[stateAndPhase.Key];
                        }
                        break;
                }
                table.Rows.Add(row);
            }

            return table;
        }
        private DataTable CreateGraphTableForCluster(string name, GraphTableDataType dataType)
        {
            DataTable table = new DataTable(name);
            DataRow row;

            table.Columns.Add(new DataColumn("Phases", typeof(string)));
            Type columnType = typeof(string);
            switch (dataType)
            {
                case GraphTableDataType.Seconds: columnType = typeof(int); break;
                case GraphTableDataType.Minutes: columnType = typeof(double); break;
                case GraphTableDataType.Percentages: columnType = typeof(double); break;
                case GraphTableDataType.Numbers: columnType = typeof(int); break;
            }

            // Add columns
            int counter = 1;
            foreach (KeyValuePair<int, Stats> curClusterAndStat in calculatedData.clusterAndStats)
            {
                table.Columns.Add(new DataColumn($"{curClusterAndStat.Key}cl", columnType));
                counter++;
            }

            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                row = table.NewRow();
                row["Phases"] = stateAndPhase.Value;
                switch (dataType)
                {
                    case GraphTableDataType.Seconds:
                        foreach (KeyValuePair<int, Stats> curClusterAndStat in calculatedData.clusterAndStats)
                        {
                            row[curClusterAndStat.Key] = curClusterAndStat.Value.StateTimes[stateAndPhase.Key];
                        }
                        break;
                    case GraphTableDataType.Minutes:
                        foreach (KeyValuePair<int, Stats> curClusterAndStat in calculatedData.clusterAndStats)
                        {
                            row[curClusterAndStat.Key] = Math.Round((double)curClusterAndStat.Value.StateTimes[stateAndPhase.Key] / 60, 2);
                        }
                        break;
                    case GraphTableDataType.Percentages:
                        foreach (KeyValuePair<int, Stats> curClusterAndStat in calculatedData.clusterAndStats)
                        {
                            row[curClusterAndStat.Key] = curClusterAndStat.Value.StatePercentages[stateAndPhase.Key];
                        }
                        break;
                    case GraphTableDataType.Numbers:
                        foreach (KeyValuePair<int, Stats> curClusterAndStat in calculatedData.clusterAndStats)
                        {
                            row[curClusterAndStat.Key] = curClusterAndStat.Value.StateNumber[stateAndPhase.Key];
                        }
                        break;
                }
                table.Rows.Add(row);
            }

            return table;
        }
        private DataTable CreateFrequencyTable(string name, Dictionary<int, SortedList<int, int>> stateFrequencies, bool isTotal = false)
        {
            DataTable table = new DataTable(name);
            DataRow row;

            // Add columns based on states
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                table.Columns.Add(new DataColumn($"{stateAndPhase.Value.Substring(0, 1)} time", typeof(int)));
                table.Columns.Add(new DataColumn($"{stateAndPhase.Value.Substring(0, 1)} frequency", typeof(int)));
            }

            // Add data from dictionary to table
            // Find largest dictionary to iterate with
            int max = 0;
            foreach (KeyValuePair<int, SortedList<int, int>> stateTimeFrequency in stateFrequencies)
            {
                if (stateTimeFrequency.Value.Count > max) { max = stateTimeFrequency.Value.Count; }
            }

            int time;
            int frequency;
            SortedList<int, int> current;
            // Iterate with largest dictionary (let's say Wakefulness has more variety than others, its dictionary will be larger)
            for (int i = 0; i < max; i++)
            {
                row = table.NewRow();

                foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
                {
                    current = stateFrequencies[stateAndPhase.Key];
                    // Since we are going with largest dictionary there will be cases that index is out of range for smaller ones
                    // in that case we just set time and frequency to 0, which we will convert into blank/null during excel export
                    if (i < current.Count)
                    {
                        time = current.ElementAt(i).Key;
                        frequency = current.ElementAt(i).Value;
                    }
                    else
                    {
                        time = 0;
                        frequency = 0;
                    }
                    row[$"{stateAndPhase.Value.Substring(0, 1)} time"] = time;
                    row[$"{stateAndPhase.Value.Substring(0, 1)} frequency"] = frequency;
                }

                table.Rows.Add(row);
            }

            return table;
        }
        private DataTable CreateCustomFrequencyTable(string name, Dictionary<int, Dictionary<string, int>> stateCustomFrequencies, bool isTotal = false)
        {
            DataTable table = new DataTable(name);
            DataRow row;

            // Add columns based on states
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                table.Columns.Add(new DataColumn($"{stateAndPhase.Value.Substring(0, 1)} range", typeof(string)));
                table.Columns.Add(new DataColumn($"{stateAndPhase.Value.Substring(0, 1)} frequency", typeof(int)));
            }

            int max = 0;
            foreach (KeyValuePair<int, Dictionary<string, int>> stateTimeFrequency in stateCustomFrequencies)
            {
                if (stateTimeFrequency.Value.Count > max) { max = stateTimeFrequency.Value.Count; }
            }

            string range;
            int frequency;
            Dictionary<string, int> current;
            for (int i = 0; i < max; i++)
            {
                row = table.NewRow();

                foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
                {
                    current = stateCustomFrequencies[stateAndPhase.Key];
                    range = current.ElementAt(i).Key;
                    frequency = current.ElementAt(i).Value;
                    row[$"{stateAndPhase.Value.Substring(0, 1)} range"] = range;
                    row[$"{stateAndPhase.Value.Substring(0, 1)} frequency"] = frequency;
                }

                table.Rows.Add(row);
            }

            return table;
        }

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
