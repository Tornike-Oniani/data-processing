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
        public Dictionary<int, List<int>> GetErrorRowIndexes(List<TimeSpan> sleepTimes, out string errorLog)
        {
            Dictionary<int, List<int>> result = new Dictionary<int, List<int>>();
            string log = "Behavior errors:\n";
            string corruptedLogFull = "";
            string overlapLogFull = "";
            string corruptedLog = "";
            string overlapLog = "";

            List<int> corrupts;
            List<int> overlaps;
            List<int> indexes;
            for (int i = 3; i <= 7; i++)
            {
                corrupts = GetCorruptedIntervalIndexes(i, out corruptedLog);
                corruptedLogFull += corruptedLog;
                overlaps = GetOverlapIntervalIndexes(i, sleepTimes, out overlapLog);
                overlapLogFull += overlapLog;
                indexes = corrupts.Concat(overlaps).Distinct().ToList();
                if (indexes.Count > 0)
                {
                    result.Add(i, indexes);
                }
            }

            log += corruptedLogFull + overlapLogFull;
            errorLog = log;
            return result;
        }
        #endregion

        #region Private helpers
        // There should be no overlap between behaviour intervals, because animal can be doing one activity at a time
        private List<int> GetOverlapIntervalIndexes(int behaviour, List<TimeSpan> sleepTimes, out string errorLog)
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
                    log += "\t- " + curInterval.From + "-" + curInterval.Till + " no overlap\n";
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
        #endregion
    }
}
