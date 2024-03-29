﻿using DataProcessing.Utils;
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
        public Dictionary<int, double> StatePercentages { get; set; } = new Dictionary<int, double>();
        public Dictionary<int, int> StateNumber { get; set; } = new Dictionary<int, int>();
        public Dictionary<SpecificCriteria, int> SpecificTimes = new Dictionary<SpecificCriteria, int>(new SpecificCriteriaComparer());
        public Dictionary<SpecificCriteria, int> SpecificNumbers = new Dictionary<SpecificCriteria, int>(new SpecificCriteriaComparer());

        public void CalculatePercentages()
        {
            foreach (KeyValuePair<int, int> entry in StateTimes)
            {
                double percentage = Math.Round((double)entry.Value * 100 / TotalTime, 2);
                StatePercentages.Add(entry.Key, percentage);
            }
        }
    }
}
