using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Models
{
    class Stats
    {
        public int TotalTime { get; set; }
        public Dictionary<int, int> StateTimes { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, double> TimePercentages { get; set; } = new Dictionary<int, double>();
        public Dictionary<int, int> StateNumber { get; set; } = new Dictionary<int, int>();
        // Refactor this bullshit
        public Dictionary<int, int> SpecificCrietriaStates { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> SpecificStateTimes { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> SpecificTimeNumbers { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> SpecificStateTimesAbove { get; set; } = new Dictionary<int, int>();
        public Dictionary<int, int> SpecificStateNumbersAbove { get; set; } = new Dictionary<int, int>();

        public void CalculatePercentages()
        {
            // Normal percentages
            foreach (KeyValuePair<int, int> entry in StateTimes)
            {
                double percentage = Math.Round((double)entry.Value * 100 / TotalTime, 2);
                TimePercentages.Add(entry.Key, percentage);
            }
        }
    }
}
