using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    internal class Behaviours
    {
        #region Private attributes
        private readonly List<Tuple<int, TimeInterval>> behaviourTimeIntervals;

        #endregion

        #region Constructors
        public Behaviours()
        {
            behaviourTimeIntervals = new List<Tuple<int, TimeInterval>>();
        }
        #endregion

        #region Public methods
        public void AddTimeIntervalOfBehaviour(TimeInterval interval, int behaviour)
        {
            behaviourTimeIntervals.Add(new Tuple<int, TimeInterval>(behaviour, interval));
        }
        public List<Tuple<int, TimeInterval>> GetSortedBehaviorTimeIntervals()
        {
            return behaviourTimeIntervals.OrderBy(bi => bi.Item2.From).ToList();
        }
        public List<Tuple<int, TimeInterval>> GetIntervalsWithinInvterval(TimeInterval interval)
        {
            List<Tuple<int, TimeInterval>> intervals = GetSortedBehaviorTimeIntervals();

            //List<Tuple<int, TimeInterval>> result = new List<Tuple<int, TimeInterval>>();
            //foreach (Tuple<int, TimeInterval> curBI in intervals)
            //{
            //    if (curBI.Item2.From >= interval.From && curBI.Item2.Till <= interval.Till)
            //    {
            //        result.Add(curBI);
            //    }
            //}

            //return result;
            return intervals
                .Where(bi => bi.Item2.From >= interval.From && bi.Item2.Till <= interval.Till)
                .ToList();
        }
        public int Count()
        {
            return behaviourTimeIntervals.Count;
        }
        public Dictionary<int, List<int>> GetErrorRowIndexes(List<TimeSpan> sleepTimes, List<TimeSpan> wakefulnessTimes, out string errorLog)
        {
            Dictionary<int, List<int>> result = new Dictionary<int, List<int>>();
            string log = "Behavior errors:\n";
            string corruptedLog = "";
            string voidLog = "";
            string duplicateLog = "";
            string overlapLog = "";
            string corruptedLogFull = "";
            string voidLogFull = "";
            string duplicateLogFull = "";
            string overlapLogFull = "";

            List<int> corrupts;
            List<int> voids;
            List<int> duplicates;
            List<int> overlaps;
            List<int> indexes;
            for (int i = 3; i <= 7; i++)
            {
                corrupts = GetCorruptedIntervalIndexes(i, out corruptedLog);
                corruptedLogFull += corruptedLog;
                voids = GetVoidIntervalIndexes(i, sleepTimes, wakefulnessTimes, out voidLog);
                voidLogFull += voidLog;
                duplicates = GetDuplicateIntervalIndexes(i, out duplicateLog);
                duplicateLogFull += duplicateLog;
                overlaps = GetOverLapIntervalIndexes(i, out overlapLog);
                overlapLogFull += overlapLog;
                indexes = corrupts.Concat(voids).Concat(duplicates).Concat(overlaps).Distinct().ToList();
                if (indexes.Count > 0)
                {
                    result.Add(i, indexes);
                }
            }

            log += corruptedLogFull + voidLogFull + overlapLogFull + duplicateLogFull;
            errorLog = log;
            return result;
        }
        #endregion

        #region Private helpers
        // The end of each behavior should be a start of another or sleep AND each start of the behavior should be another's end or wakefulness (except for the first one). There should be no blanks.
        private List<int> GetVoidIntervalIndexes(int behaviour, List<TimeSpan> sleepTimes, List<TimeSpan> wakefulnessTimes, out string errorLog)
        {
            List<int> result = new List<int>();
            string log = "";

            TimeInterval[] intervals = behaviourTimeIntervals
                                .Where(bi => bi.Item1 == behaviour)
                                .Select(bi => bi.Item2)
                                .ToArray();
            TimeInterval curInterval;
            for (int i = 0; i < intervals.Length; i++)
            {
                curInterval = intervals[i];
                // If the end of the interval doesn't match the start of any behavior or sleep then its an error
                if (!behaviourTimeIntervals.Any(bi => bi.Item2.From == curInterval.Till) &&
                    !sleepTimes.Any(w => w == curInterval.Till))
                {
                    log += "\t- " + curInterval.From + "-" + curInterval.Till + " void interval at end\n";
                    result.Add(i + 1);
                    continue;
                }
                // If the start of the interval doesn't match the end of any behavior or wakefulness its an error (except for the first one which starts at 00:00:00)
                if (curInterval.From != new TimeSpan(0, 0, 0) && !behaviourTimeIntervals.Any(bi => bi.Item2.Till == curInterval.From) && !wakefulnessTimes.Any(w => w == curInterval.From))
                {
                    log += "\t- " + curInterval.From + "-" + curInterval.Till + " void interval at start\n";
                    result.Add(i + 1);
                }
            }

            errorLog = log;
            return result;
        }
        private List<int> GetCorruptedIntervalIndexes(int behaviour, out string errorLog)
        {
            List<int> result = new List<int>();
            string log = "";

            TimeInterval[] intervals = behaviourTimeIntervals
                                            .Where(bi => bi.Item1 == behaviour)
                                            .Select(bi => bi.Item2)
                                            .ToArray();

            TimeInterval curInterval;
            for (int i = 0; i < intervals.Length - 1; i++)
            {
                curInterval = intervals[i];
                if (!curInterval.IsCorrect())
                {
                    // We add i + 1 because this will be row indexes in excel sheet where indexing starts with 1 and not 0
                    log += "\t- " + curInterval.From + "-" + curInterval.Till + " incorrect interval\n";
                    result.Add(i + 1);
                }
            }

            errorLog = log;
            return result;
        }
        private List<int> GetDuplicateIntervalIndexes(int behaviour, out string errorLog)
        {
            List<int> result = new List<int>();
            string log = "";

            TimeInterval[] intervals = behaviourTimeIntervals
                                .Where(bi => bi.Item1 == behaviour)
                                .Select(bi => bi.Item2)
                                .ToArray();

            TimeInterval curInterval;
            for (int i = 0; i < intervals.Length; i++)
            {
                curInterval = intervals[i];
                // If an interval is used more than once its an error
                if (behaviourTimeIntervals.Count(bi => bi.Item2.From == curInterval.From && bi.Item2.Till == curInterval.Till) > 1)
                {
                    log += "\t- " + curInterval.From + "-" + curInterval.Till + " used more than once (duplicate)\n";
                    result.Add(i + 1);
                }
            }

            errorLog = log;
            return result;
        }
        private List<int> GetOverLapIntervalIndexes(int behaviour, out string errorLog)
        {
            List<int> result = new List<int>();
            string log = "";

            TimeInterval[] intervals = behaviourTimeIntervals
                                .Where(bi => bi.Item1 == behaviour)
                                .Select(bi => bi.Item2)
                                .ToArray();

            TimeInterval curInterval;
            for (int i = 0; i < intervals.Length; i++)
            {
                curInterval = intervals[i];
                // If an interval overlaps with any other its an error
                if (CheckOverlap(behaviourTimeIntervals.Select(bi => bi.Item2).ToList(), curInterval))
                {
                    log += "\t- " + curInterval.From + "-" + curInterval.Till + " overlaps with other intervals\n";
                    result.Add(i + 1);
                }
            }

            errorLog = log;
            return result;
        }

        private bool CheckOverlap(List<TimeInterval> intervals, TimeInterval sample)
        {
            foreach (TimeInterval interval in intervals)
            {
                // Since we are checking interval overalp in a list which also contains said interval we want to skip the check on itself. We are interested if it overlaps with other ones.
                if (sample.From == interval.From && sample.Till == interval.Till)
                {
                    continue;
                }
                if (sample.From < interval.Till && sample.Till > interval.From)
                {
                    return true; // Overlapping interval found
                }
            }
            return false; // No overlapping intervals found
        }
        #endregion
    }
}
