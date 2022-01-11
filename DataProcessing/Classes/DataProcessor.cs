using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;

namespace DataProcessing.Classes
{
    enum GraphTableDataType
    {
        Seconds,
        Minutes,
        Percentages,
        Numbers
    }

    class DataProcessor
    {
        // Private attributes
        private ExportOptions options;
        private List<TimeStamp> timeStamps;
        private Dictionary<int, string> stateAndPhases = new Dictionary<int, string>();
        private Stats totalStats;
        private Dictionary<int, Stats> hourAndStats = new Dictionary<int, Stats>();
        // State frequencies total + each hour
        private List<Dictionary<int, SortedList<int, int>>> hourStateFrequencies = new List<Dictionary<int, SortedList<int, int>>>();
        private List<int> hourRowIndexes = new List<int>();
        private List<Tuple<int, string>> hourRowIndexesTime = new List<Tuple<int, string>>();
        private List<DataTableInfo> statTableCollection = new List<DataTableInfo>();
        private List<DataTableInfo> graphTableCollection = new List<DataTableInfo>();
        private List<Tuple<int, int>> duplicatedTimes = new List<Tuple<int, int>>();
        private List<DataTableInfo> frequenciesCollection = new List<DataTableInfo>();
        private int timeBeforeFirstSleep = 0;
        private int timeBeforeFirstParadoxicalSleep = 0;

        // Constructor
        public DataProcessor(List<TimeStamp> timeStamps, ExportOptions options)
        {
            this.timeStamps = timeStamps;
            this.options = options;
            List<int> states = timeStamps.Where(sample => sample.State != 0).Select(sample => sample.State).Distinct().ToList();
            states.Sort();

            if (states.Count > options.MaxStates) { throw new Exception($"File contains more than {options.MaxStates} states!"); }

            CreatePhases();
        }

        // Public methods
        public void Calculate(List<TimeStamp> nonMarkedRecords)
        {
            // Create duplicated timestamps for graph
            int previous = timeStamps[0].TimeDifferenceInSeconds;
            duplicatedTimes.Add(new Tuple<int, int>(previous, timeStamps[1].State));
            for (int i = 1; i < timeStamps.Count; i++)
            {
                duplicatedTimes.Add(new Tuple<int, int>(timeStamps[i].TimeDifferenceInSeconds + previous, timeStamps[i].State));
                if (i < timeStamps.Count - 1)
                {
                    duplicatedTimes.Add(new Tuple<int, int>(timeStamps[i].TimeDifferenceInSeconds + previous, timeStamps[i + 1].State));
                }
                previous = previous + timeStamps[i].TimeDifferenceInSeconds;
            }

            // Calculate total
            totalStats = CalculateStats(timeStamps, true);

            // Calculate per hour
            int time = 0;
            int currentHour = 0;

            // Hourly frequencies
            Dictionary<int, SortedList<int, int>> totalFrequencies = new Dictionary<int, SortedList<int, int>>();
            Dictionary<int, SortedList<int, int>> hourFrequencies = new Dictionary<int, SortedList<int, int>>();

            foreach (KeyValuePair<int, string> stateAndPhase in stateAndPhases)
            {
                // Create time and frequency dictionary for the state
                totalFrequencies.Add(stateAndPhase.Key, new SortedList<int, int>());
                hourFrequencies.Add(stateAndPhase.Key, new SortedList<int, int>());
            }

            // Add total here so it will be on top of hourly frequencies
            hourStateFrequencies.Add(totalFrequencies);

            // Calculate total frequencies with non marked original timestamps
            for (int i = 1; i < nonMarkedRecords.Count; i++)
            {
                TimeStamp currentTimeStamp = nonMarkedRecords[i];

                // We don't want program added timestamps (marker and hour marks) to be added to total
                if (!currentTimeStamp.IsTimeMarked && !currentTimeStamp.IsMarker)
                    AddFrequencyToCollection(totalFrequencies, currentTimeStamp);
            }

            List<TimeStamp> hourRegion = new List<TimeStamp>();
            // Latency
            bool foundFirstSleep = false;
            bool foundFirstParadoxicalSleep = false;
            int lastHourIndex = 0;
            for (int i = 0; i < timeStamps.Count; i++)
            {
                TimeStamp currentTimeStamp = timeStamps[i];
                time += currentTimeStamp.TimeDifferenceInSeconds;

                // Calculate time before first sleep and paradoxical sleep (Latency)
                if (!foundFirstSleep)
                {
                    timeBeforeFirstSleep += currentTimeStamp.TimeDifferenceInSeconds;
                    if (currentTimeStamp.State == 2) { foundFirstSleep = true; }
                }
                if (!foundFirstParadoxicalSleep)
                {
                    timeBeforeFirstParadoxicalSleep += currentTimeStamp.TimeDifferenceInSeconds;
                    if (currentTimeStamp.State == 1) { foundFirstParadoxicalSleep = true; }
                }

                if (time > options.TimeMark * 3600) { throw new Exception("Invalid hour marks"); }

                hourRegion.Add(currentTimeStamp);
                lastHourIndex = i + 1;

                // Add frequencies, first timestamp doesn't have a state so we skip it
                if (i > 0)
                {
                    AddFrequencyToCollection(hourFrequencies, currentTimeStamp);
                }

                if (time == options.TimeMark * 3600)
                {
                    currentHour++;
                    hourRowIndexes.Add(i + 1);
                    hourRowIndexesTime.Add(new Tuple<int, string>(i + 1, $"Hour {currentHour * options.TimeMark}"));
                    hourAndStats.Add(currentHour, CalculateStats(hourRegion, false));
                    hourStateFrequencies.Add(hourFrequencies);

                    time = 0;
                    hourRegion.Clear();

                    // Reset hour frequency collection
                    hourFrequencies = new Dictionary<int, SortedList<int, int>>();
                    foreach (KeyValuePair<int, string> stateAndPhase in stateAndPhases)
                    {
                        // Create time and frequency dictionary for the state
                        hourFrequencies.Add(stateAndPhase.Key, new SortedList<int, int>());
                    }
                }
            }

            // Do last part (might be less than marked time)
            if (hourRegion.Count == 0) { return; }
            currentHour++;
            hourAndStats.Add(currentHour, CalculateStats(hourRegion, false));
            hourStateFrequencies.Add(hourFrequencies);
            hourRowIndexes.Add(lastHourIndex);
            hourRowIndexesTime.Add(new Tuple<int, string>(lastHourIndex, getTimeForStats(hourAndStats.Last().Value.TotalTime)));
        }
        public List<DataTableInfo> CreateStatTables()
        {
            CreateStatTable("Total", totalStats, true);

            int counter = 1;
            foreach (KeyValuePair<int, Stats> hourAndStat in hourAndStats)
            {
                if (counter == hourAndStats.Count && hourAndStat.Value.TotalTime % 3600 != 0)
                {
                    CreateStatTable(getTimeForStats(hourAndStat.Value.TotalTime), hourAndStat.Value, false);
                    continue;
                }
                CreateStatTable($"hour {hourAndStat.Key * options.TimeMark}", hourAndStat.Value, false);
                counter++;
            }

            return statTableCollection;
        }
        public List<DataTableInfo> CreateGraphTables()
        {
            CreateGraphTable("Percentages %", GraphTableDataType.Percentages);
            CreateGraphTable("Minutes", GraphTableDataType.Minutes);
            CreateGraphTable("Seconds", GraphTableDataType.Seconds);
            CreateGraphTable("Numbers", GraphTableDataType.Numbers);
            return graphTableCollection;
        }
        public List<Tuple<int, int>> getDuplicatedTimes() { return duplicatedTimes; }
        public List<int> getHourRowIndexes() { return hourRowIndexes; }
        public List<Tuple<int, string>> getHourRowIndexesTime() { return hourRowIndexesTime; }
        public List<DataTableInfo> CreateFrequencyTables()
        {
            int hour = 0;
            foreach (Dictionary<int, SortedList<int, int>> stateTimeFrequency in hourStateFrequencies)
            {
                // First item will be total not hourly
                if (hour == 0)
                {
                    CreateFrequencyTable("Total", stateTimeFrequency, true);
                    hour++;
                    continue;
                }

                CreateFrequencyTable(hourRowIndexesTime[hour - 1].Item2, stateTimeFrequency);
                hour++;
            }

            return frequenciesCollection;
        }
        public DataTableInfo CreateLatencyTable()
        {
            DataTableInfo tableInfo = new DataTableInfo();
            DataTable table = new DataTable("Latency");
            tableInfo.Table = table;

            table.Columns.Add(new DataColumn("First sleep", typeof(int)));
            table.Columns.Add(new DataColumn("First PS", typeof(int)));

            DataRow row = table.NewRow();
            row["First sleep"] = timeBeforeFirstSleep;
            row["First PS"] = timeBeforeFirstParadoxicalSleep;
            table.Rows.Add(row);

            tableInfo.HeaderIndexes = new Tuple<int, int>(1, 1);
            tableInfo.PhasesIndexes = new Tuple<int, int>(2, 2);

            return tableInfo;
        }

        // Private helpers
        private void CreateStatTable(string name, Stats stats, bool isTotal)
        {
            DataTableInfo tableInfo = new DataTableInfo();
            DataTable table = new DataTable(name);
            tableInfo.Table = table;
            DataRow row;

            table.Columns.Add(new DataColumn("Phases", typeof(string)));
            table.Columns.Add(new DataColumn("sec", typeof(int)));
            table.Columns.Add(new DataColumn("min", typeof(double)));
            table.Columns.Add(new DataColumn("%", typeof(double)));
            table.Columns.Add(new DataColumn("num", typeof(int)));
            tableInfo.HeaderIndexes = new Tuple<int, int>(0, 5);

            foreach (KeyValuePair<int, string> stateAndPhase in stateAndPhases)
            {
                row = table.NewRow();
                row["Phases"] = stateAndPhase.Value;
                row["sec"] = stats.StateTimes[stateAndPhase.Key];
                row["min"] = Math.Round((double)stats.StateTimes[stateAndPhase.Key] / 60, 2);
                row["%"] = stats.StatePercentages[stateAndPhase.Key];
                row["num"] = stats.StateNumber[stateAndPhase.Key];
                table.Rows.Add(row);
            }
            int phaseCount = stateAndPhases.Count;

            if (isTotal)
            {
                row = table.NewRow();
                row["Phases"] = "Total";
                row["sec"] = stats.TotalTime;
                row["min"] = Math.Round((double)stats.TotalTime / 60, 2);
                table.Rows.Add(row);
                tableInfo.IsTotal = true;
                phaseCount++;
            }
            tableInfo.PhasesIndexes = new Tuple<int, int>(1, phaseCount);

            foreach (SpecificCriteria criteria in options.Criterias)
            {
                // Skip nonexistent criterias
                if (criteria.Value == null) { continue; }

                row = table.NewRow();
                row["Phases"] = $"{stateAndPhases[criteria.State]} {criteria.GetOperandValue()} {criteria.Value}";
                row["sec"] = stats.SpecificTimes[criteria];
                row["min"] = Math.Round((double)stats.SpecificTimes[criteria] / 60, 2);
                row["num"] = stats.SpecificNumbers[criteria];
                table.Rows.Add(row);
            }
            tableInfo.CriteriaPhases = new Tuple<int, int>(phaseCount, options.Criterias.Where(x => x.Value != null).Count());

            statTableCollection.Add(tableInfo);
        }
        private void CreateGraphTable(string name, GraphTableDataType dataType)
        {
            DataTableInfo tableInfo = new DataTableInfo();
            DataTable table = new DataTable(name);
            tableInfo.Table = table;
            graphTableCollection.Add(tableInfo);
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

            // Columns
            int counter = 1;
            foreach (KeyValuePair<int, Stats> hourAndStat in hourAndStats)
            {
                if (counter == hourAndStats.Count && hourAndStat.Value.TotalTime % 3600 != 0)
                {
                    table.Columns.Add(new DataColumn(getTimeForGraph(hourAndStat.Value.TotalTime), columnType));
                    continue;
                }
                table.Columns.Add(new DataColumn($"{hourAndStat.Key * options.TimeMark}hr", columnType));
                counter++;
            }
            // +1 for title
            tableInfo.HeaderIndexes = new Tuple<int, int>(0, hourAndStats.Count + 1);

            foreach (KeyValuePair<int, string> stateAndPhase in stateAndPhases)
            {
                row = table.NewRow();
                row["Phases"] = stateAndPhase.Value;
                switch (dataType)
                {
                    case GraphTableDataType.Seconds:
                        foreach (KeyValuePair<int, Stats> hourAndStat in hourAndStats)
                        {
                            row[hourAndStat.Key] = hourAndStat.Value.StateTimes[stateAndPhase.Key];
                        }
                        break;
                    case GraphTableDataType.Minutes:
                        foreach (KeyValuePair<int, Stats> hourAndStat in hourAndStats)
                        {
                            row[hourAndStat.Key] = Math.Round((double)hourAndStat.Value.StateTimes[stateAndPhase.Key] / 60, 2);
                        }
                        break;
                    case GraphTableDataType.Percentages:
                        foreach (KeyValuePair<int, Stats> hourAndStat in hourAndStats)
                        {
                            row[hourAndStat.Key] = hourAndStat.Value.StatePercentages[stateAndPhase.Key];
                        }
                        break;
                    case GraphTableDataType.Numbers:
                        foreach (KeyValuePair<int, Stats> hourAndStat in hourAndStats)
                        {
                            row[hourAndStat.Key] = hourAndStat.Value.StateNumber[stateAndPhase.Key];
                        }
                        break;
                }
                table.Rows.Add(row);
            }
            tableInfo.PhasesIndexes = new Tuple<int, int>(1, stateAndPhases.Count);
        }
        private void CreatePhases()
        {
            if (options.MaxStates == 3)
            {
                stateAndPhases = new Dictionary<int, string>();
                stateAndPhases.Add(3, "Wakefulness");
                stateAndPhases.Add(2, "Sleep");
                stateAndPhases.Add(1, "Paradoxical sleep");
                
                

            }
            else if (options.MaxStates == 4)
            {
                stateAndPhases = new Dictionary<int, string>();
                stateAndPhases.Add(4, "Wakefulness");
                stateAndPhases.Add(3, "Light sleep");
                stateAndPhases.Add(2, "Deep sleep");
                stateAndPhases.Add(1, "Paradoxical sleep");
            }
            else
            {
                throw new Exception($"Max states can be either 3 or 4");
            }
        }
        private Stats CalculateStats(List<TimeStamp> region, bool forTotal)
        {
            Stats result = new Stats();
            result.TotalTime = region.Sum((sample) => sample.TimeDifferenceInSeconds);

            foreach (int state in stateAndPhases.Keys)
            {
                result.StateTimes.Add(state, calculateStateTime(region, state));
                result.StateNumber.Add(state, calculateStateNumber(region, state, forTotal));
            }
            result.CalculatePercentages();

            foreach (SpecificCriteria criteria in options.Criterias)
            {
                // Skip nonexistent crietrias
                if (criteria.Value == null) { continue; }

                result.SpecificTimes.Add(criteria, calculateStateCriteriaTime(region, criteria));
                result.SpecificNumbers.Add(criteria, calculateStateCriteriaNumber(region, criteria, forTotal));
            }

            return result;
        }
        private void CreateFrequencyTable(string name, Dictionary<int, SortedList<int, int>> stateFrequencies, bool isTotal = false)
        {
            DataTableInfo tableInfo = new DataTableInfo();
            DataTable table = new DataTable(name);
            tableInfo.Table = table;
            DataRow row;

            // Add columns based on states
            foreach (KeyValuePair<int, string> stateAndPhase in stateAndPhases)
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

                foreach (KeyValuePair<int, string> stateAndPhase in stateAndPhases)
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

            tableInfo.HeaderIndexes = new Tuple<int, int>(1, 1);
            tableInfo.PhasesIndexes = new Tuple<int, int>(1, table.Columns.Count);
            tableInfo.IsTotal = isTotal;

            frequenciesCollection.Add(tableInfo);
        }

        private int calculateStateTime(List<TimeStamp> region, int state)
        {
            return region.Where((sample) => sample.State == state).Select((sample) => sample.TimeDifferenceInSeconds).Sum();
        }
        private int calculateStateNumber(List<TimeStamp> region, int state, bool forTotal = false)
        {
            if (forTotal)
            {
                return region.Count(sample => sample.State == state && !sample.IsMarker && !sample.IsTimeMarked);
            }
            return region.Count(sample => sample.State == state);
        }
        private int calculateStateCriteriaTime(List<TimeStamp> samples, SpecificCriteria criteria)
        {
            if (criteria.Operand == "Below")
            {
                return samples.Where((sample) => sample.State == criteria.State && sample.TimeDifferenceInSeconds <= criteria.Value).Select((sample) => sample.TimeDifferenceInSeconds).Sum();
            }

            return samples.Where((sample) => sample.State == criteria.State && sample.TimeDifferenceInSeconds >= criteria.Value).Select((sample) => sample.TimeDifferenceInSeconds).Sum();

        }
        private int calculateStateCriteriaNumber(List<TimeStamp> samples, SpecificCriteria criteria, bool forTotal = false)
        {
            if (criteria.Operand == "Below")
            {
                if (forTotal)
                {
                    return samples.Count(sample => sample.State == criteria.State && !sample.IsMarker && !sample.IsTimeMarked && sample.TimeDifferenceInSeconds <= criteria.Value);
                }
                return samples.Count(sample => sample.State == criteria.State && sample.TimeDifferenceInSeconds <= criteria.Value);
            }

            if (forTotal)
            {
                return samples.Count(sample => sample.State == criteria.State && !sample.IsMarker && !sample.IsTimeMarked && sample.TimeDifferenceInSeconds >= criteria.Value);
            }
            return samples.Count(sample => sample.State == criteria.State && sample.TimeDifferenceInSeconds >= criteria.Value);
        }
        private string getCriteriaLabel(SpecificCriteria criteria)
        {
            return $"{stateAndPhases[criteria.State]} {criteria.GetOperandValue()} {criteria.Value}";
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
            // This is weird it gives minutes as float hours
            double result = Math.Round((double)seconds / 3600, 2);
            if (result < 1) { return $"Last {result % 1}hr"; }
            return $"Last {Math.Truncate(result)}hr {result % 1}min";
        }
        private void AddFrequencyToCollection(Dictionary<int, SortedList<int, int>> collection, TimeStamp timeStamp)
        {
            // If time is already enetered increment frequency
            if (collection[timeStamp.State].ContainsKey(timeStamp.TimeDifferenceInSeconds))
            {
                collection[timeStamp.State][timeStamp.TimeDifferenceInSeconds] += 1;
            }
            // Otherwise add it with frequency 1
            else
            {
                collection[timeStamp.State].Add(timeStamp.TimeDifferenceInSeconds, 1);
            }
        }
    }
}
