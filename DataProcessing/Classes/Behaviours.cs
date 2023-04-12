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
        // There should be no overlap between behaviour intervals, because animal can be doing one activity at a time
        public bool CheckBehaviourOverlap()
        {
            bool isOverlap = false;
            List<Tuple<int, TimeInterval>> intervals = GetSortedBehaviorTimeIntervals();
            return isOverlap;

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
        #endregion
    }
}
