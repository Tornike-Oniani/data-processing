using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Models
{
    class Stats
    {
        public int TotalTime { get; set; }
        public Dictionary<int, int> StateTimes { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, double> StatePercentages { get; set; } = new Dictionary<int, double>();
        public Dictionary<int, int> StateNumber { get; set; } = new Dictionary<int, int>();
        public Dictionary<SpecificCriteria, int> SpecificTimes = new Dictionary<SpecificCriteria, int>(new SpecificCriteriaComparer());
        public Dictionary<SpecificCriteria, int> SpecificNumbers = new Dictionary<SpecificCriteria, int>(new SpecificCriteriaComparer());

        public void CalculatePercentages()
        {
            foreach (KeyValuePair<int, int> entry in StateTimes)
            {
                double percentage;
                if (TotalTime == 0)
                {
                    percentage = 0;
                }
                else 
                { 
                    percentage = Math.Round((double)entry.Value * 100 / TotalTime, 2);
                }
                StatePercentages.Add(entry.Key, percentage);
            }
        }
        public void CalculateBehavioralPercentages(int[] states, int wakefulnessTime)
        {
            foreach (int state in states)
            {
                double percentage = Math.Round((double)StateTimes[state] / wakefulnessTime, 2);
                StatePercentages.Add(state, percentage);
            }
        }
    }
}
