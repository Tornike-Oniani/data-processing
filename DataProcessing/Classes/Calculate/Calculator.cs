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
        public int calculateStateTime(List<TimeStamp> region, int state)
        {
            return region.Where((sample) => sample.State == state).Select((sample) => sample.TimeDifferenceInSeconds).Sum();
        }
        public int calculateStateNumber(List<TimeStamp> region, int state)
        {
            return region.Count(sample => sample.State == state);
        }
        public int calculateStateCriteriaTime(List<TimeStamp> samples, SpecificCriteria criteria)
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
        public int calculateStateCriteriaNumber(List<TimeStamp> samples, SpecificCriteria criteria)
        {
            if (criteria.Operand == "Below")
            {
                return samples.Count(sample => sample.State == criteria.State && sample.TimeDifferenceInSeconds <= criteria.Value);
            }

            return samples.Count(sample => sample.State == criteria.State && sample.TimeDifferenceInSeconds >= criteria.Value);
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
        public Dictionary<int, SortedList<int, int>> calculateTotalFrequencies(List<TimeStamp> nonMarkedRegion, int[] states)
        {
            Dictionary<int, SortedList<int, int>> result = new Dictionary<int, SortedList<int, int>>();

            foreach (int state in states)
            {
                // Initialize sorted list for each phase
                result.Add(state, new SortedList<int, int>());
            }

            // Calculate total frequencies with non marked original timestamps
            for (int i = 1; i < nonMarkedRegion.Count; i++)
            {
                TimeStamp currentTimeStamp = nonMarkedRegion[i];

                // We don't want program added timestamps (marker and hour marks) to be added to total
                if (!currentTimeStamp.IsTimeMarked && !currentTimeStamp.IsMarker)
                {
                    AddFrequencyToCollection(result, currentTimeStamp);
                }
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

            for (int i = 1; i < region.Count; i++)
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
        #endregion
    }
}
