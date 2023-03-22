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
    }
}
