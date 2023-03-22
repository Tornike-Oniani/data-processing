using DataProcessing.Classes.Calculate;
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
        private readonly CalculatedData calculatedData;
        // Exported options selected by user
        private readonly CalculationOptions options;
        private readonly int _criteriaNumber;
        // Table decorator for coloring
        private readonly TableDecorator decorator;

        // Constructor
        public TableCreator(CalculationOptions calcOptions, CalculatedData data)
        {
            // Init
            calculatedData = data;
            options = calcOptions;
            _criteriaNumber = options.Criterias.Count(c => c.Value != null);
            decorator = new TableDecorator(options.MaxStates);
        }

        // Public table creators
        public ExcelTable CreateRawDataTable()
        {
            int rowCount = options.MarkedTimeStamps.Count;
            int colCount = 5;
            object[,] data = new object[rowCount, colCount];

            // Fill in data
            int rowIndex = 0;
            foreach (TimeStamp timeStamp in options.MarkedTimeStamps)
            {
                // We convert time into string other wise setting range value in excel won't work
                data[rowIndex, 0] = timeStamp.Time.ToString();
                data[rowIndex, 1] = timeStamp.TimeDifference.ToString();
                data[rowIndex, 2] = timeStamp.TimeDifferenceInDouble;
                data[rowIndex, 3] = timeStamp.TimeDifferenceInSeconds;
                data[rowIndex, 4] = timeStamp.State;
                rowIndex++;
            }

            // Decorate collection and return it
            return decorator.DecorateRawData(data, options.MarkedTimeStamps, options.TimeMarkInSeconds);
        }
        public ExcelTable CreateLatencyTable()
        {
            // Title + header + data
            int rowCount = 3;
            // If we have 3 states we have first sleep and first PS
            // but if we have only 2 then we only want first sleep
            int colCount = options.MaxStates - 1;
            object[,] data = new object[rowCount, colCount];

            // Set title
            data[0, 0] = "Latency";

            // Set header
            data[1, 0] = "First sleep";
            if (options.MaxStates == 3)
            {
                data[1, 1] = "First PS";
            }

            // Fill in data
            data[2, 0] = calculatedData.timeBeforeFirstSleep;
            if (options.MaxStates == 3)
            {
                data[2, 1] = calculatedData.timeBeforeFirstParadoxicalSleep;
            }

            // Decorate table and return it
            return decorator.DecorateLatencyTable(data);
        }
        public List<ExcelTable> CreateStatTables()
        {
            List<ExcelTable> tables = new List<ExcelTable>();

            // Create a total stat table and add it on top
            tables.Add(CreateStatTable("Total", calculatedData.totalStats, true));

            // Add table for each hour mark
            int counter = 1;
            string tableName;
            foreach (KeyValuePair<int, Stats> hourAndStat in calculatedData.hourAndStats)
            {
                tableName = $"episode {counter}";
                tables.Add(CreateStatTable(tableName, hourAndStat.Value, false));
                counter++;
            }

            // Decorate collection and return it
            return tables;
        }
        public List<ExcelTable> CreateGraphTables()
        {
            List<ExcelTable> tables = new List<ExcelTable>();

            // Create graph tables
            tables.Add(CreateGraphTable("Percentages %", GraphTableDataType.Percentages, false, true));
            tables.Add(CreateGraphTable("Minutes", GraphTableDataType.Minutes, false));
            tables.Add(CreateGraphTable("Seconds", GraphTableDataType.Seconds, false));
            tables.Add(CreateGraphTable("Numbers", GraphTableDataType.Numbers, false));

            // Decorate collection and return it
            return tables;
        }
        public ExcelTable CreateDuplicatesTable()
        {
            int rowCount = calculatedData.duplicatedTimes.Count;
            int colCount = 2;
            object[,] data = new object[rowCount, colCount];

            // Fill in data
            int rowIndex = 0;
            foreach (Tuple<int, int> item in calculatedData.duplicatedTimes)
            {
                data[rowIndex, 0] = item.Item1;
                data[rowIndex, 1] = item.Item2;
                rowIndex++;
            }

            // Decorate collection and return it
            return decorator.DecorateDuplicatesTable(data);
        }
        public List<ExcelTable> CreateFrequencyTables()
        {
            List<ExcelTable> tables = new List<ExcelTable>();

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

                tables.Add(CreateFrequencyTable($"episode {hour}", stateTimeFrequency));
                hour++;
            }

            return tables;
        }
        public List<ExcelTable> CreateCustomFrequencyTables()
        {
            // If user didn't provide any custom frequency
            if (options.FrequencyRanges.Count == 0) { return null; }

            List<ExcelTable> tables = new List<ExcelTable>();

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

                tables.Add(CreateCustomFrequencyTable($"episode {hour}", stateTimeFrequency));
                hour++;
            }

            return tables;
        }
        public ExcelTable CreateClusterDataTable()
        {
            int rowCount = options.NonMarkedTimeStamps.Count;
            int colCount = 2;
            object[,] data = new object[rowCount, colCount];

            // Fill in data
            int rowIndex = 0;
            foreach (TimeStamp timeStamp in options.NonMarkedTimeStamps)
            {
                data[rowIndex, 0] = timeStamp.TimeDifferenceInSeconds;
                data[rowIndex, 1] = timeStamp.State;
                rowIndex++;
            }

            // Decorate collection and return it
            return decorator.DecorateClusterDataTable(data, options.NonMarkedTimeStamps, options.ClusterSeparationTimeInSeconds);
        }
        public List<ExcelTable> CreateGraphTablesForClusters()
        {
            List<ExcelTable> tables = new List<ExcelTable>();

            // Create graph tables
            tables.Add(CreateGraphTable("Percentages %", GraphTableDataType.Percentages, true, true));
            tables.Add(CreateGraphTable("Minutes", GraphTableDataType.Minutes, true));
            tables.Add(CreateGraphTable("Seconds", GraphTableDataType.Seconds, true));
            tables.Add(CreateGraphTable("Numbers", GraphTableDataType.Numbers, true));

            // Decorate collection and return it
            return tables;
        }

        // Single table helper creators
        private ExcelTable CreateStatTable(string name, Stats stats, bool isTotal)
        {
            // Header + all the phases + optional criterias + total row if its total
            int rowCount =
                calculatedData.stateAndPhases.Count +
                options.Criterias.Count(c => c.Value != null) +
                1 +
                (isTotal ? 1 : 0);
            object[,] data = new object[rowCount, 5];

            // Set title
            data[0, 0] = name;

            // Add columns
            data[0, 1] = "sec";
            data[0, 2] = "min";
            data[0, 3] = "%";
            data[0, 4] = "num";

            // Fill in essential data
            // We start from 1 because 0 is set to header
            int rowIndex = 1;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                data[rowIndex, 0] = stateAndPhase.Value;
                data[rowIndex, 1] = stats.StateTimes[stateAndPhase.Key];
                data[rowIndex, 2] = Math.Round((double)stats.StateTimes[stateAndPhase.Key] / 60, 2);
                data[rowIndex, 3] = stats.StatePercentages[stateAndPhase.Key];
                data[rowIndex, 4] = stats.StateNumber[stateAndPhase.Key];
                rowIndex++;
            }

            // If table is total add one more row for summed up stats
            if (isTotal)
            {
                data[rowIndex, 0] = "Total";
                data[rowIndex, 1] = stats.TotalTime;
                data[rowIndex, 2] = Math.Round((double)stats.TotalTime / 60, 2);
                rowIndex++;
            }

            // Fill in data for specific criterias if any was set
            foreach (SpecificCriteria criteria in options.Criterias)
            {
                // Skip nonexistent criterias
                if (criteria.Value == null) { continue; }

                data[rowIndex, 0] = $"{calculatedData.stateAndPhases[criteria.State]} {criteria.GetOperandValue()} {criteria.Value}";
                data[rowIndex, 1] = stats.SpecificTimes[criteria];
                data[rowIndex, 2] = Math.Round((double)stats.SpecificTimes[criteria] / 60, 2);
                data[rowIndex, 4] = stats.SpecificNumbers[criteria];
                rowIndex++;
            }

            if (isTotal) { return decorator.DecorateStatTableTotal(data, _criteriaNumber); }
            return decorator.DecorateStatTable(data, _criteriaNumber);
        }
        // Division can be either hourAndStats or clusterAndStats
        private ExcelTable CreateGraphTable(string name, GraphTableDataType dataType, bool isCluster, bool hasChart = false)
        {
            Dictionary<int, Stats> division = isCluster ? calculatedData.clusterAndStats : calculatedData.hourAndStats;

            // Header + phases
            int rowCount = calculatedData.stateAndPhases.Count + 1;
            // Phases + each hour mark
            int colCount = division.Count + 1;
            object[,] data = new object[rowCount, colCount];

            // Set title
            data[0, 0] = name;

            // Add columns
            // We start from 1 because 0 is set to title
            int colIndex = 1;
            foreach (KeyValuePair<int, Stats> hourAndStat in division)
            {
                data[0, colIndex] = $"{colIndex}" + (isCluster ? "cl" : "ep");
                colIndex++;
            }

            // Fill in data
            // We start from 1 because 0 is set to header
            int rowIndex = 1;
            colIndex = 0;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                data[rowIndex, colIndex] = stateAndPhase.Value;
                colIndex++;
                switch (dataType)
                {
                    case GraphTableDataType.Seconds:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            data[rowIndex, colIndex] = hourAndStat.Value.StateTimes[stateAndPhase.Key];
                            colIndex++;
                        }
                        break;
                    case GraphTableDataType.Minutes:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            data[rowIndex, colIndex] = Math.Round((double)hourAndStat.Value.StateTimes[stateAndPhase.Key] / 60, 2);
                            colIndex++;
                        }
                        break;
                    case GraphTableDataType.Percentages:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            data[rowIndex, colIndex] = hourAndStat.Value.StatePercentages[stateAndPhase.Key];
                            colIndex++;
                        }
                        break;
                    case GraphTableDataType.Numbers:
                        foreach (KeyValuePair<int, Stats> hourAndStat in division)
                        {
                            data[rowIndex, colIndex] = hourAndStat.Value.StateNumber[stateAndPhase.Key];
                            colIndex++;
                        }
                        break;
                }
                colIndex = 0;
                rowIndex++;
            }

            return decorator.DecorateGraphTable(data, hasChart);
        }
        private ExcelTable CreateFrequencyTable(string name, Dictionary<int, SortedList<int, int>> stateFrequencies, bool isTotal = false)
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
            object[,] data = new object[rowCount, colCount];

            // Set title
            data[0, 0] = name;

            // Add columns based on states
            int colIndex = 0;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                data[1, colIndex] = $"{stateAndPhase.Value.Substring(0, 1)} time";
                data[1, colIndex + 1] = $"{stateAndPhase.Value.Substring(0, 1)} freq";
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
                        data[rowIndex, colIndex] = time;
                        data[rowIndex, colIndex + 1] = frequency;
                    }
                    else
                    {
                        // Since we can't pass null to datatable column we have to use DBNull.Value instead
                        data[rowIndex, colIndex] = null;
                        data[rowIndex, colIndex + 1] = null;
                    }
                    colIndex += 2;
                }
                rowIndex++;
                colIndex = 0;
            }

            return decorator.DecorateFrequencyTable(data, isTotal);
        }
        private ExcelTable CreateCustomFrequencyTable(string name, Dictionary<int, Dictionary<string, int>> stateCustomFrequencies, bool isTotal = false)
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
            object[,] data = new object[rowCount, colCount];

            // Set title
            data[0, 0] = name;

            // Add columns based on states
            data[1, 0] = "Ranges";
            int colIndex = 1;
            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                data[1, colIndex] = $"{stateAndPhase.Value.Substring(0, 1)} freq";
                colIndex++;
            }

            string range;
            int frequency;
            Dictionary<string, int> current;
            int rowIndex = 2;
            colIndex = 1;
            for (int i = 0; i < max; i++)
            {
                foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
                {
                    current = stateCustomFrequencies[stateAndPhase.Key];
                    range = current.ElementAt(i).Key;
                    frequency = current.ElementAt(i).Value;
                    data[rowIndex, 0] = range;
                    data[rowIndex, colIndex] = frequency;
                    colIndex++;
                }
                colIndex = 1;
                rowIndex++;
            }

            return decorator.DecorateCustomFrequencyTable(data, options.FrequencyRanges.Count, isTotal);
        }
    }
}
