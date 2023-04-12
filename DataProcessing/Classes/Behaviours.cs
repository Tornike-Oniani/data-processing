using System;
using System.Collections.Generic;
using System.Linq;
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
        public Dictionary<int, List<int>> GetErrorRowIndexes()
        {
            Dictionary<int, List<int>> result = new Dictionary<int, List<int>>();

            for (int i = 3; i <= 7; i++)
            {

            }
        }
        #endregion

        #region Private helpers
        // There should be no overlap between behaviour intervals, because animal can be doing one activity at a time
        private List<int> CheckBehaviourOverlap(List<TimeSpan> sleepTimes, int behaviour)
        {
            List<int> result = new List<int>();

            TimeInterval[] intervals = behaviourTimeIntervals
                                .Where(bi => bi.Item1 == behaviour)
                                .Select(bi => bi.Item2)
                                .ToArray();
            TimeInterval curInterval;
            for (int i = 0; i < intervals.Length; i++)
            {
                curInterval = intervals[i];
                // If the end of the interval doesn't match the start of any behavior or sleep then its an error
                if (behaviourTimeIntervals.Any(bi => bi.Item2.From == curInterval.Till) &&
                    sleepTimes.Any(w => w == curInterval.Till))
                {
                    result.Add(i + 1);
                }
            }

            return result;
        }
        private List<int> GetCorruptedIntervalIndexes(int behaviour)
        {
            List<int> result = new List<int>();
            TimeInterval[] intervals = behaviourTimeIntervals
                                            .Where(bi => bi.Item1 == behaviour)
                                            .Select(bi => bi.Item2)
                                            .ToArray();

            for (int i = 0; i < intervals.Length - 1; i++)
            {
                if (!intervals[i].IsCorrect())
                {
                    // We add i + 1 because this will be row indexes in excel sheet where indexing starts with 1 and not 0
                    result.Add(i + 1);
                }
            }

            return result;
        }
        #endregion
    }
}
