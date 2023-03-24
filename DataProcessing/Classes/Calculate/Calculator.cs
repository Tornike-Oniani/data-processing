using DataProcessing.Models;
using DataProcessing.Utils;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes.Calculate
{
    internal class Calculator
    {
        public Stats CalculateStats(List<TimeStamp> region, int[] states, List<SpecificCriteria> criterias)
        {
            Stats result = new Stats();
            result.TotalTime = region.Sum((sample) => sample.TimeDifferenceInSeconds);

            foreach (int state in states)
            {
                result.StateTimes.Add(state,calculateStateTime(region, state));
                result.StateNumber.Add(state, calculateStateNumber(region, state));
            }
            result.CalculatePercentages();

            foreach (SpecificCriteria criteria in criterias)
            {
                // Skip nonexistent crietrias
                if (criteria.Value == null) { continue; }

                result.SpecificTimes.Add(criteria, calculateStateCriteriaTime(region, criteria));
                result.SpecificNumbers.Add(criteria, calculateStateCriteriaNumber(region, criteria));
            }

            return result;
        }
        public List<List<TimeStamp>> CreateClusters(List<TimeStamp> records, int clusterSeparationTime, int wakefulnessState)
        {
            List<List<TimeStamp>> result = new List<List<TimeStamp>>();

            List<TimeStamp> cluster = new List<TimeStamp>();
            TimeStamp curTimeStamp;
            for (int i = 1; i < records.Count; i++)
            {
                curTimeStamp = records[i];
                // If we found end of the cluster add it to clusters list
                if (curTimeStamp.TimeDifferenceInSeconds >= clusterSeparationTime && curTimeStamp.State == wakefulnessState)
                {
                    // If we found end of cluster but it doesn't contain any timestamps we don't want to calculate stats for it. This can happen if recording starts with long wakefulness - firs record will be 0-0 and then essentialy a cluster end. We don't want to include blank clusters like this
                    if (cluster.Count == 0) { continue; }

                    result.Add(cluster);
                    cluster = new List<TimeStamp>();
                    continue;
                }
                cluster.Add(records[i]);
            }

            // If we have remainder in the last cluster add it to (The list won't always end with wakefulness that is more than cluster separation time)
            if (cluster.Count > 0)
            {
                result.Add(cluster);
            }

            return result;
        }
        public Dictionary<int, Stats> CreateStatsForClusters(List<TimeStamp> region, int clusterSeparationTime, int wakefulnessState, int[] states, List<SpecificCriteria> criterias)
        {
            Dictionary<int, Stats> result = new Dictionary<int, Stats>();

            List<TimeStamp> clusterRegion = new List<TimeStamp>();
            TimeStamp curTimeStamp;
            int curClusterNumber = 1;
            for (int i = 1; i < region.Count; i++)
            {
                curTimeStamp = region[i];

                // If we found end of the cluster calculate its stats and add it to dictionary
                if (curTimeStamp.TimeDifferenceInSeconds >= clusterSeparationTime && curTimeStamp.State == wakefulnessState)
                {
                    // If we found end of cluster but it doesn't contain any timestamps we don't want to calculate stats for it. This can happen if recording starts with long wakefulness - firs record will be 0-0 and then essentialy a cluster end. We don't want to include blank clusters like this
                    if (clusterRegion.Count == 0) { continue; }

                    result.Add(curClusterNumber, CalculateStats(clusterRegion, states, criterias));
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
                result.Add(curClusterNumber, CalculateStats(clusterRegion, states, criterias));
            }

            return result;
        }
        // We use this list for ladder-graph
        public List<Tuple<int, int>> generateDuplicatedTimeStamps(List<TimeStamp> timeStamps)
        {
            List<Tuple<int, int>> result = new List<Tuple<int, int>>();

            int previous = timeStamps[0].TimeDifferenceInSeconds;
            result.Add(new Tuple<int, int>(previous, timeStamps[1].State));
            for (int i = 1; i < timeStamps.Count; i++)
            {
                result.Add(new Tuple<int, int>(timeStamps[i].TimeDifferenceInSeconds + previous, timeStamps[i].State));
                if (i < timeStamps.Count - 1)
                {
                    result.Add(new Tuple<int, int>(timeStamps[i].TimeDifferenceInSeconds + previous, timeStamps[i + 1].State));
                }
                previous = previous + timeStamps[i].TimeDifferenceInSeconds;
            }

            return result;
        }
        public Dictionary<int, SortedList<int, int>> calculateFrequencies(List<TimeStamp> region, int[] states)
        {
            Dictionary<int, SortedList<int, int>> result = new Dictionary<int, SortedList<int, int>>();

            foreach (int state in states)
            {
                // Initialize sorted list for each phase
                result.Add(state, new SortedList<int, int>());
            }

            // Calculate total frequencies with non marked original timestamps
            for (int i = region[0].State == 0 ? 1 : 0 ; i < region.Count; i++)
            {
                TimeStamp currentTimeStamp = region[i];
                AddFrequencyToCollection(result, currentTimeStamp);
            }

            return result;
        }
        public Dictionary<int, Dictionary<string, int>> calculateFrequencyRanges(List<TimeStamp> region, int[] states, Dictionary<string, int[]> frequencyRanges)
        {
            Dictionary<int, Dictionary<string, int>> result = new Dictionary<int, Dictionary<string, int>>();

            // Initialize collection for each state
            foreach (int state in states)
            {
                result.Add(state, new Dictionary<string, int>());
            }

            // Initialize ranges as keys
            foreach (KeyValuePair<int, Dictionary<string, int>> stateRange in result)
            {
                foreach (KeyValuePair<string, int[]> range in frequencyRanges)
                {
                    stateRange.Value.Add(range.Key, 0);
                }
            }

            for (int i = 0; i < region.Count; i++)
            {
                TimeStamp currentTimeStamp = region[i];

                // Find fitting range for current timestamp
                foreach (KeyValuePair<string, int[]> range in frequencyRanges)
                {
                    if (
                        currentTimeStamp.TimeDifferenceInSeconds >= range.Value[0] &&
                        currentTimeStamp.TimeDifferenceInSeconds <= range.Value[1])
                    {
                        result[currentTimeStamp.State][range.Key] += 1;
                    }
                }
            }

            return result;
        }
        public int calculateStateLatency(List<TimeStamp> region, int state)
        {
            return region.TakeWhile(t => t.State != state).Sum(t => t.TimeDifferenceInSeconds);
        }

        #region Private helpers
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
        private int calculateStateTime(List<TimeStamp> region, int state)
        {
            return region.Where((sample) => sample.State == state).Select((sample) => sample.TimeDifferenceInSeconds).Sum();
        }
        private int calculateStateNumber(List<TimeStamp> region, int state)
        {
            return region.Count(sample => sample.State == state);
        }
        private int calculateStateCriteriaTime(List<TimeStamp> samples, SpecificCriteria criteria)
        {
            if (criteria.Operand == "Below")
            {
                return samples.Where((sample) => sample.State == criteria.State && sample.TimeDifferenceInSeconds <= criteria.Value).Select((sample) => sample.TimeDifferenceInSeconds).Sum();
            }

            return samples
                .Where((sample) => sample.State == criteria.State && sample.TimeDifferenceInSeconds >= criteria.Value)
                .Select((sample) => sample.TimeDifferenceInSeconds)
                .Sum();

        }
        private int calculateStateCriteriaNumber(List<TimeStamp> samples, SpecificCriteria criteria)
        {
            if (criteria.Operand == "Below")
            {
                return samples.Count(sample => sample.State == criteria.State && sample.TimeDifferenceInSeconds <= criteria.Value);
            }

            return samples.Count(sample => sample.State == criteria.State && sample.TimeDifferenceInSeconds >= criteria.Value);
        }
        #endregion
    }
}
