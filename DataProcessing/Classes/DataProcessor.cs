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

    /// <summary>
    /// Takes initial data and calculates necessary stats for exporting
    /// </summary>
    internal class DataProcessor
    {
        // Private attributes
        private ExportOptions options;
        private List<TimeStamp> timeStamps;
        private List<TimeStamp> nonMarkedTimeStamps;
        private CalculatedData calculatedData = new CalculatedData();

        // Constructor
        public DataProcessor(List<TimeStamp> timeStamps, List<TimeStamp> nonMarkedTimeStamps, ExportOptions options)
        {
            // Init
            this.timeStamps = timeStamps;
            this.nonMarkedTimeStamps = nonMarkedTimeStamps;
            this.options = options;

            // Extract all distinct states from excel file
            List<int> states = timeStamps.Where(sample => sample.State != 0).Select(sample => sample.State).Distinct().ToList();
            states.Sort();

            // If number of extracted states doesn't match number of selected states throw error.
            if (states.Count > options.MaxStates) { throw new Exception($"File contains more than {options.MaxStates} states!"); }

            // Map state numbers to phase strings (e.g 1 - PS, 2 - Sleep, 3 - Wakefulness)
            CreatePhases();
        }

        // Public methods
        public CalculatedData Calculate()
        {
            // Create duplicated timestamps for graph
            int previous = timeStamps[0].TimeDifferenceInSeconds;
            calculatedData.duplicatedTimes.Add(new Tuple<int, int>(previous, timeStamps[1].State));
            for (int i = 1; i < timeStamps.Count; i++)
            {
                calculatedData.duplicatedTimes.Add(new Tuple<int, int>(timeStamps[i].TimeDifferenceInSeconds + previous, timeStamps[i].State));
                if (i < timeStamps.Count - 1)
                {
                    calculatedData.duplicatedTimes.Add(new Tuple<int, int>(timeStamps[i].TimeDifferenceInSeconds + previous, timeStamps[i + 1].State));
                }
                previous = previous + timeStamps[i].TimeDifferenceInSeconds;
            }

            // Calculate total
            calculatedData.totalStats = CalculateStats(timeStamps, true);

            // Calculate per hour
            int time = 0;
            int currentHour = 0;

            // Hourly frequencies
            Dictionary<int, SortedList<int, int>> totalFrequencies = new Dictionary<int, SortedList<int, int>>();
            Dictionary<int, SortedList<int, int>> hourFrequencies = new Dictionary<int, SortedList<int, int>>();

            // Hourly custom frequncies
            Dictionary<int, Dictionary<string, int>> totalCustomFrequencies = new Dictionary<int, Dictionary<string, int>>();
            Dictionary<int, Dictionary<string, int>> hourCustomFrequencies = new Dictionary<int, Dictionary<string, int>>();

            foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
            {
                // Create time and frequency dictionary for the state
                totalFrequencies.Add(stateAndPhase.Key, new SortedList<int, int>());
                hourFrequencies.Add(stateAndPhase.Key, new SortedList<int, int>());

                // Create the same for customs
                totalCustomFrequencies.Add(stateAndPhase.Key, new Dictionary<string, int>());
                hourCustomFrequencies.Add(stateAndPhase.Key, new Dictionary<string, int>());
            }

            // Add ranges to each state for custom frequencies
            foreach (KeyValuePair<int, Dictionary<string, int>> stateRange in totalCustomFrequencies)
            {
                foreach (KeyValuePair<string, int[]> range in options.customFrequencyRanges)
                {
                    stateRange.Value.Add(range.Key, 0);
                }
            }
            foreach (KeyValuePair<int, Dictionary<string, int>> stateRange in hourCustomFrequencies)
            {
                foreach (KeyValuePair<string, int[]> range in options.customFrequencyRanges)
                {
                    stateRange.Value.Add(range.Key, 0);
                }
            }

            // Add total here so it will be on top of hourly frequencies
            calculatedData.hourStateFrequencies.Add(totalFrequencies);
            calculatedData.hourStateCustomFrequencies.Add(totalCustomFrequencies);

            // Calculate total frequencies with non marked original timestamps
            for (int i = 1; i < nonMarkedTimeStamps.Count; i++)
            {
                TimeStamp currentTimeStamp = nonMarkedTimeStamps[i];

                // We don't want program added timestamps (marker and hour marks) to be added to total
                if (!currentTimeStamp.IsTimeMarked && !currentTimeStamp.IsMarker)
                    AddFrequencyToCollection(totalFrequencies, currentTimeStamp);

                // Find fitting range for current timestamp
                foreach (KeyValuePair<string, int[]> range in options.customFrequencyRanges)
                {
                    if (
                        currentTimeStamp.TimeDifferenceInSeconds >= range.Value[0] && 
                        currentTimeStamp.TimeDifferenceInSeconds <= range.Value[1])
                    {
                        totalCustomFrequencies[currentTimeStamp.State][range.Key] += 1;
                    }
                }
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
                    if (currentTimeStamp.State == 2)
                        foundFirstSleep = true;
                    else
                        calculatedData.timeBeforeFirstSleep += currentTimeStamp.TimeDifferenceInSeconds;
                }
                if (!foundFirstParadoxicalSleep)
                {
                    if (currentTimeStamp.State == 1)
                        foundFirstParadoxicalSleep = true;
                    else
                        calculatedData.timeBeforeFirstParadoxicalSleep += currentTimeStamp.TimeDifferenceInSeconds;
                }

                if (time > options.TimeMark) { throw new Exception("Invalid hour marks"); }

                hourRegion.Add(currentTimeStamp);
                lastHourIndex = i + 1;

                // Add frequencies, first timestamp doesn't have a state so we skip it
                if (i > 0)
                {
                    AddFrequencyToCollection(hourFrequencies, currentTimeStamp);
                    AddCustomFrequencyToCollection(hourCustomFrequencies, currentTimeStamp);
                }

                if (time == options.TimeMark)
                {
                    currentHour++;
                    //hourRowIndexes.Add(i + 1);
                    //hourRowIndexesTime.Add(new Tuple<int, string>(i + 1, $"Hour {currentHour * options.TimeMark}"));
                    calculatedData.hourAndStats.Add(currentHour, CalculateStats(hourRegion, false));
                    calculatedData.hourStateFrequencies.Add(hourFrequencies);
                    calculatedData.hourStateCustomFrequencies.Add(hourCustomFrequencies);

                    time = 0;
                    hourRegion.Clear();

                    // Reset hour frequency collection
                    hourFrequencies = new Dictionary<int, SortedList<int, int>>();
                    foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
                    {
                        // Create time and frequency dictionary for the state
                        hourFrequencies.Add(stateAndPhase.Key, new SortedList<int, int>());
                    }

                    // Reset hour custom frequency collection
                    hourCustomFrequencies = new Dictionary<int, Dictionary<string, int>>();
                    foreach (KeyValuePair<int, string> stateAndPhase in calculatedData.stateAndPhases)
                    {
                        hourCustomFrequencies.Add(stateAndPhase.Key, new Dictionary<string, int>());
                    }
                    foreach (KeyValuePair<int, Dictionary<string, int>> stateRange in hourCustomFrequencies)
                    {
                        foreach (KeyValuePair<string, int[]> range in options.customFrequencyRanges)
                        {
                            stateRange.Value.Add(range.Key, 0);
                        }
                    }
                }
            }

            // Do last part (might be less than marked time)
            if (hourRegion.Count != 0) 
            {
                currentHour++;
                calculatedData.hourAndStats.Add(currentHour, CalculateStats(hourRegion, false));
                calculatedData.hourStateFrequencies.Add(hourFrequencies);
                calculatedData.hourStateCustomFrequencies.Add(hourCustomFrequencies);
                //hourRowIndexes.Add(lastHourIndex);
                //hourRowIndexesTime.Add(new Tuple<int, string>(lastHourIndex, getTimeForStats(hourAndStats.Last().Value.TotalTime)));
            }

            // Calculate stats for clusters
            if (options.ClusterSeparationTimeInSeconds > 0)
            {
                CreateStatsForClusters();
            }

            return calculatedData;
        }


        // Private helpers
        private void CreateStatsForClusters()
        {
            List<TimeStamp> clusterRegion = new List<TimeStamp>();
            TimeStamp curTimeStamp;
            int curClusterNumber = 1;
            for (int i = 1; i < nonMarkedTimeStamps.Count; i++)
            {
                curTimeStamp = nonMarkedTimeStamps[i];

                // If we found end of the cluster calculate its stats and add it to dictionary
                if (curTimeStamp.TimeDifferenceInSeconds >= options.ClusterSeparationTimeInSeconds && curTimeStamp.State == 3)
                {
                    // If we found end of cluster but it doesn't contain any timestamps we don't want to calculate stats for it
                    // This can happen if recording starts with long wakefulness - firs record will be 0-0 and then essentialy a cluster end
                    // We don't want to include blank clusters like this
                    if (clusterRegion.Count == 0) { continue; }
                    calculatedData.clusterAndStats.Add(curClusterNumber, CalculateStats(clusterRegion, false));
                    curClusterNumber++;
                    // After the stats of current cluster was calculated we reset it for the next one
                    clusterRegion = new List<TimeStamp>();
                    continue;
                }

                clusterRegion.Add(curTimeStamp);
            }

            // If we have remainder in the last cluster calculate its stats too (The list won't always end with wakefulness that is more than cluster separation time)
            if (clusterRegion.Count > 0)
            {
                calculatedData.clusterAndStats.Add(curClusterNumber, CalculateStats(clusterRegion, false));
            }
        }
        private void CreatePhases()
        {
            if (options.MaxStates == 2)
            {
                calculatedData.stateAndPhases = new Dictionary<int, string>();
                calculatedData.stateAndPhases.Add(2, "Wakefulness");
                calculatedData.stateAndPhases.Add(1, "Sleep");
            }
            else if (options.MaxStates == 3)
            {
                calculatedData.stateAndPhases = new Dictionary<int, string>();
                calculatedData.stateAndPhases.Add(3, "Wakefulness");
                calculatedData.stateAndPhases.Add(2, "Sleep");
                calculatedData.stateAndPhases.Add(1, "Paradoxical sleep");
            }
            else if (options.MaxStates == 4)
            {
                calculatedData.stateAndPhases = new Dictionary<int, string>();
                calculatedData.stateAndPhases.Add(4, "Wakefulness");
                calculatedData.stateAndPhases.Add(3, "Light sleep");
                calculatedData.stateAndPhases.Add(2, "Deep sleep");
                calculatedData.stateAndPhases.Add(1, "Paradoxical sleep");
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

            foreach (int state in calculatedData.stateAndPhases.Keys)
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
        private void AddCustomFrequencyToCollection(Dictionary<int, Dictionary<string, int>> collection, TimeStamp timeStamp)
        {
            // Find fitting range for current timestamp
            foreach (KeyValuePair<string, int[]> range in options.customFrequencyRanges)
            {
                if (
                    timeStamp.TimeDifferenceInSeconds >= range.Value[0] &&
                    timeStamp.TimeDifferenceInSeconds <= range.Value[1])
                {
                    collection[timeStamp.State][range.Key] += 1;
                }
            }
        }
    }
}
