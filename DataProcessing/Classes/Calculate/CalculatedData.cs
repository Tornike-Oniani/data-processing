using DataProcessing.Constants;
using DataProcessing.Models;
using Microsoft.Office.Interop.Excel;
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
        #region Public properties
        public Dictionary<int, string> stateAndPhases { get; set; }
        public Stats totalStats { get; set; }
        public Dictionary<int, Stats> hourAndStats { get; set; }
        public Dictionary<int, Stats> clusterAndStats { get; set; }
        // State frequencies total + each hour
        public List<Dictionary<int, SortedList<int, int>>> stateFrequencies { get; set; }
        public List<Dictionary<int, Dictionary<string, int>>> stateFrequencyRanges{ get; set; }
        public int timeBeforeFirstSleep { get; set; }
        public int timeBeforeFirstParadoxicalSleep { get; set; }
        public List<Tuple<int, int>> duplicatedTimes { get; set; }
        #endregion

        #region Constructors
        // Constructor
        public CalculatedData()
        {
            // Init
            stateAndPhases = new Dictionary<int, string>();
            hourAndStats = new Dictionary<int, Stats>();
            clusterAndStats = new Dictionary<int, Stats>();
            stateFrequencies = new List<Dictionary<int, SortedList<int, int>>>();
            stateFrequencyRanges= new List<Dictionary<int, Dictionary<string, int>>>();
            timeBeforeFirstSleep = 0;
            timeBeforeFirstParadoxicalSleep = 0;
            duplicatedTimes = new List<Tuple<int, int>>();
        }
        #endregion

        #region Public methods
        public void CreatePhases(int maxStates)
        {
            if (maxStates == 2)
            {
                stateAndPhases = RecordingType.GetTwoStatesDictionary();
            }
            else if (maxStates == 3)
            {
                stateAndPhases = RecordingType.GetThreeStatesDictionary();
            }
            else if (maxStates == 7)
            {
                stateAndPhases = RecordingType.GetTwoStatesWithBehaviorDictionary();
            }
            else if (maxStates == 4)
            {
                stateAndPhases = new Dictionary<int, string>();
                stateAndPhases.Add(4, "Wakefulness");
                stateAndPhases.Add(3, "Light sleep");
                stateAndPhases.Add(2, "Deep sleep");
                stateAndPhases.Add(1, "Paradoxical sleep");
            }
            else
            {
                throw new Exception("Max states can be either 2 or 3");
            }
        }
        public int[] GetStates()
        {
            return stateAndPhases.Keys.ToArray();
        }
        public void AddFrequency(Dictionary<int, SortedList<int, int>> frequency)
        {
            stateFrequencies.Add(frequency);
        }
        public void AddFrequencyRange(Dictionary<int, Dictionary<string, int>> frequencyRange)
        {
            stateFrequencyRanges.Add(frequencyRange);
        }
        #endregion
    }
}
