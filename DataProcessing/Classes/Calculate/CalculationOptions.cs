using DataProcessing.Constants;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataProcessing.Classes.Calculate
{
    internal class CalculationOptions
    {
        #region Public properties
        public int TimeMarkInSeconds { get; set; }
        public int MaxStates { get; set; }
        public Dictionary<int, string> StateAndPhases { get; set; }
        public List<TimeStamp> MarkedTimeStamps { get; set; }
        public List<TimeStamp> NonMarkedTimeStamps { get; set; }
        // In behavior recording type we want a separated list that shows just sleep and wakefulness
        public List<TimeStamp> NonMarkedNormalizedTimeStamps { get; set; }
        public List<TimeStamp> MarkedNormalizedTimeStamps { get; set; }
        public Dictionary<string, int[]> FrequencyRanges { get; set; }
        public int ClusterSeparationTimeInSeconds { get; set; }
        public List<SpecificCriteria> Criterias { get; set; }
        #endregion

        #region Constructors
        public CalculationOptions(List<TimeStamp> region, UserSelectedOptions options)
        {
            TimeMarkInSeconds = ConvertTimeMarkToSeconds(options.SelectedTimeMark);
            FrequencyRanges = options.FrequencyRanges;
            Criterias = options.Criterias;
            ClusterSeparationTimeInSeconds = options.ClusterSparationTime * 60;
            MaxStates = RecordingType.MaxStates[options.SelectedRecordingType];
            MapStateToPhases(MaxStates);
            // We have switched assignment of normalized and regular!!!!!!!
            if (MaxStates == 8)
            {
                NonMarkedTimeStamps = CloneAndNormalizeTimeStamps(region);
                NonMarkedNormalizedTimeStamps =  CloneTimeStamps(region);
                MarkedTimeStamps = AddTimeMarksToSamples(CloneAndNormalizeTimeStamps(region));
                MarkedNormalizedTimeStamps = AddTimeMarksToSamples(region);
            }
            else
            {
                NonMarkedTimeStamps = CloneTimeStamps(region);
                NonMarkedNormalizedTimeStamps = CloneAndNormalizeTimeStamps(region);
                MarkedNormalizedTimeStamps = AddTimeMarksToSamples(CloneAndNormalizeTimeStamps(region));
                MarkedTimeStamps = AddTimeMarksToSamples(region);
            }
            CalculateSamples(NonMarkedTimeStamps);
            CalculateSamples(MarkedTimeStamps);
            CalculateSamples(NonMarkedNormalizedTimeStamps);
            CalculateSamples(MarkedNormalizedTimeStamps);
        }
        #endregion

        #region Public methods
        public int GetState(string phase)
        {
            return StateAndPhases.FirstOrDefault(sap => sap.Value == phase).Key;
        }
        public int[] GetAllStates()
        {
            return StateAndPhases.Keys.ToArray();
        }
        #endregion

        #region Private helpers
        private List<TimeStamp> CloneAndNormalizeTimeStamps(List<TimeStamp> records)
        {
            List<TimeStamp> result = new List<TimeStamp>() { records[0] };

            TimeStamp wakeTimeStamp = new TimeStamp();
            bool cloneHappened = false;
            TimeStamp currTimeStamp;
            for (int i = 1; i < records.Count; i++)
            {
                currTimeStamp = records[i];
                if (currTimeStamp.State == 1)
                {
                    if (cloneHappened)
                    {
                        result.Add(wakeTimeStamp);
                        cloneHappened = false;
                    }
                    result.Add(currTimeStamp);
                }
                else if (currTimeStamp.State > 1)
                {
                    wakeTimeStamp = currTimeStamp.Clone();
                    wakeTimeStamp.State = 2;
                    cloneHappened = true;
                }

                // If last one isn't sleeping state add it
                if (i == records.Count - 1 && currTimeStamp.State != 1)
                {
                    result.Add(wakeTimeStamp);
                }
            }

            return result;
        }
        private List<TimeStamp> AddTimeMarksToSamples(List<TimeStamp> records)
        {
            List<TimeStamp> result = new List<TimeStamp>();

            TimeSpan markCap = TimeSpan.FromSeconds(TimeMarkInSeconds);
            TimeSpan markSum = new TimeSpan(0, 0, 0);
            TimeSpan lastMarkTime = records[0].Time;

            result.Add(records[0]);
            for (int i = 1; i < records.Count; i++)
            {
                TimeSpan span = records[i].Time;
                int state = records[i].State;

                // Difference between times
                if (span > records[i - 1].Time)
                    markSum += span - records[i - 1].Time;
                else
                    markSum += span + new TimeSpan(24, 0, 0) - records[i - 1].Time;

                // If sum exceeds one hour
                if (markSum > markCap)
                {
                    TimeSpan timeMark = new TimeSpan(0, 0, 0);
                    // If several mark can be put (for example lets say sum is 1 hour and 20 minutes and our mark is 30 minuntes 
                    // we should put 2 marks in that case)
                    for (int j = 1; j <= (int)(markSum.TotalSeconds / markCap.TotalSeconds); j++)
                    {
                        timeMark = lastMarkTime + markCap >= new TimeSpan(24, 0, 0) ?
                            lastMarkTime + markCap - new TimeSpan(24, 0, 0) :
                            lastMarkTime + markCap;
                        // If file already contains timestamp that exactly the same as our generated one don't add it to avoid duplicating
                        if (timeMark != span)
                        {
                            result.Add(new TimeStamp() { Time = timeMark, State = state, IsTimeMarked = true });
                            lastMarkTime = timeMark;
                        }
                    }

                    markSum = span - timeMark;
                }

                // If sum is exactly one hour reset
                if (markSum == markCap)
                {
                    markSum = new TimeSpan(0, 0, 0);
                    lastMarkTime = span;
                }

                result.Add(new TimeStamp() { Time = span, State = state, IsMarker = records[i].IsMarker });
            }

            return result;
        }
        private void CalculateSamples(List<TimeStamp> records)
        {
            for (int i = 1; i < records.Count; i++)
            {
                records[i].CalculateStatsWhenMany(records[i - 1]);
            }
        }
        private int ConvertTimeMarkToSeconds(string timeMark)
        {
            switch (timeMark)
            {
                case "10min":
                    return 600;
                case "15min":
                    return 900;
                case "20min":
                    return 1200;
                case "30min":
                    return 1800;
                case "1hr":
                    return 3600;
                case "2hr":
                    return 7200;
                case "4hr":
                    return 14400;
                default:
                    throw new Exception("Time mark does not exists");
            }
        }
        private List<TimeStamp> CloneTimeStamps(List<TimeStamp> timeStamps)
        {
            List<TimeStamp> result = new List<TimeStamp>();
            foreach (TimeStamp timeStamp in timeStamps)
            {
                result.Add(timeStamp.Clone());
            }

            return result;
        }
        public void MapStateToPhases(int maxStates)
        {
            if (maxStates == 2)
            {
                StateAndPhases = RecordingType.GetTwoStatesDictionary();
            }
            else if (maxStates == 3)
            {
                StateAndPhases = RecordingType.GetThreeStatesDictionary();
            }
            else if (maxStates == 7)
            {
                StateAndPhases = RecordingType.GetTwoStatesWithBehaviorDictionary();
            }
            else if (maxStates == 4)
            {
                StateAndPhases = new Dictionary<int, string>()
                {
                    {4, "Wakefulness" },
                    {3, "Light sleep" },
                    {2, "Deep sleep" },
                    {1, "Paradoxical sleep" }
                };
            }
            else
            {
                throw new Exception("Max states can be either 2 or 3");
            }
        }
        #endregion
    }
}
