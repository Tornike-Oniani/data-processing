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
        public Dictionary<int, string> behaviorStateAndPhases { get; set; }
        public Stats totalStats { get; set; }
        public Stats totalBehaviorStats { get; set; }
        public Dictionary<int, Stats> hourAndStats { get; set; }
        public Dictionary<int, Stats> hourAndBehaviorStats { get; set; }
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
            totalBehaviorStats = new Stats();
            behaviorStateAndPhases = new Dictionary<int, string>();
            hourAndStats = new Dictionary<int, Stats>();
            hourAndBehaviorStats = new Dictionary<int, Stats>();
            clusterAndStats = new Dictionary<int, Stats>();
            stateFrequencies = new List<Dictionary<int, SortedList<int, int>>>();
            stateFrequencyRanges= new List<Dictionary<int, Dictionary<string, int>>>();
            timeBeforeFirstSleep = 0;
            timeBeforeFirstParadoxicalSleep = 0;
            duplicatedTimes = new List<Tuple<int, int>>();
        }
        #endregion

        #region Public methods
        public void MapStateToPhases(string recordingType)
        {
            if (recordingType == RecordingType.TwoStates)
            {
                stateAndPhases = RecordingType.GetTwoStatesDictionary();
            }
            else if (recordingType == RecordingType.ThreeStates)
            {
                stateAndPhases = RecordingType.GetThreeStatesDictionary();
            }
            else if (recordingType == RecordingType.TwoStatesWithBehavior)
            {
                stateAndPhases = RecordingType.GetTwoStatesDictionary();
                behaviorStateAndPhases = RecordingType.GetBehaviorStatesDictionary();
            }
            else
            {
                throw new Exception("Selected recording type is not supported.");
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
