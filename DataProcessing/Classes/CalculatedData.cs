using DataProcessing.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    /// <summary>
    /// Calculated data by DataProcessor
    /// </summary>
    internal class CalculatedData
    {
        public Dictionary<int, string> stateAndPhases { get; set; }
        public Stats totalStats { get; set; }
        public Dictionary<int, Stats> hourAndStats { get; set; }
        public Dictionary<int, Stats> clusterAndStats { get; set; }
        // State frequencies total + each hour
        public List<Dictionary<int, SortedList<int, int>>> hourStateFrequencies { get; set; }
        public List<Dictionary<int, Dictionary<string, int>>> hourStateCustomFrequencies { get; set; }
        public int timeBeforeFirstSleep { get; set; }
        public int timeBeforeFirstParadoxicalSleep { get; set; }
        public List<Tuple<int, int>> duplicatedTimes { get; set; }

        // Constructor
        public CalculatedData()
        {
            // Init
            stateAndPhases = new Dictionary<int, string>();
            hourAndStats = new Dictionary<int, Stats>();
            clusterAndStats = new Dictionary<int, Stats>();
            hourStateFrequencies = new List<Dictionary<int, SortedList<int, int>>>();
            hourStateCustomFrequencies = new List<Dictionary<int, Dictionary<string, int>>>();
            timeBeforeFirstSleep = 0;
            timeBeforeFirstParadoxicalSleep = 0;
            duplicatedTimes = new List<Tuple<int, int>>();
        }
    }
}
