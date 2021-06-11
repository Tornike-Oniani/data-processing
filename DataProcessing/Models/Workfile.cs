using DataProcessing.Repositories;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Models
{
    class Workfile
    {
        // Private attributes
        private List<int> states = new List<int>();
        private string MarkerStates;

        // Public properties
        public int Id { get; set; }
        public string Name { get; set; }
        public string ImportDate { get; set; }
        public Stats Stats { get; set; } = new Stats();
        public Dictionary<int, Stats> HourlyStats { get; set; } = new Dictionary<int, Stats>();
        public Dictionary<int, int[]> HourlyIndexes { get; set; } = new Dictionary<int, int[]>();
        public List<Tuple<int, int>> DuplicatedTimes { get; set; } = new List<Tuple<int, int>>();
        public Dictionary<int, string> StatesMapping { get; set; }
        public double LastSectionTime { get; set; }
        public List<int> GetMarkerStates()
        {
            List<int> result = new List<int>();
            foreach (string state in MarkerStates.Split(','))
            {
                result.Add(int.Parse(state));
            }
            return result;
        }
        public void SetMarkerStates(string markerStates) { this.MarkerStates = markerStates; }

        // Public actions
        public void CalculateStats(List<TimeStamp> calculatedSamples, ExportOptions options)
        {
            // Calculate total stats
            Stats.TotalTime = calculatedSamples.Sum((sample) => sample.TimeDifferenceInSeconds);
            states = calculatedSamples.Where(sample => sample.State != 0).Select(sample => sample.State).Distinct().ToList();
            states.Sort();

            if (states.Count > options.MaxStates) { ClearStats(); throw new Exception($"File contains more than {options.MaxStates} states!"); }

            // Map states and phases
            if (options.MaxStates == 3)
            {
                StatesMapping = new Dictionary<int, string>();
                StatesMapping.Add(1, "Paradoxical sleep");
                StatesMapping.Add(2, "Sleep");
                StatesMapping.Add(3, "Wakefulness");

            }
            else if (options.MaxStates == 4)
            {
                StatesMapping = new Dictionary<int, string>();
                StatesMapping.Add(1, "Paradoxical sleep");
                StatesMapping.Add(2, "Deep sleep");
                StatesMapping.Add(3, "Light sleep");
                StatesMapping.Add(4, "Wakefulness");
            }
            else
            {
                ClearStats();
                throw new Exception($"Max states can be either 3 or 4");
            }


            foreach (int state in StatesMapping.Keys)
            {
                Stats.StateTimes.Add(state, calculateStateTime(calculatedSamples, state));
                Stats.StateNumber.Add(state, calculateStateNumber(calculatedSamples, state, true));
            }

            // Calculate specific criteria states
            foreach (KeyValuePair<int, int> entry in options.StateAndCriteria)
            {
                //OG
                //Stats.SpecificCrietriaStates.Add(entry.Key, calculateStateCriteriaNumber(calculatedSamples, entry.Key, entry.Value));

                Stats.SpecificStateTimes.Add(entry.Key, calculateStateCriteriaTime(calculatedSamples, entry.Key, entry.Value));
                Stats.SpecificTimeNumbers.Add(entry.Key, calculateStateCriteriaNumber(calculatedSamples, entry.Key, entry.Value));
            }
            foreach (KeyValuePair<int, int> entry in options.StateAndCriteriaAbove)
            {
                // OG
                //statHourly.SpecificCrietriaStates.Add(entry.Key, calculateStateCriteriaNumber(hourSamples, entry.Key, entry.Value));

                Stats.SpecificStateNumbersAbove.Add(entry.Key, calculateStateCriteriaNumberAbove(calculatedSamples, entry.Key, entry.Value));
                Stats.SpecificStateTimesAbove.Add(entry.Key, calculateStateCriteriaTimeAbove(calculatedSamples, entry.Key, entry.Value));
            }

            Stats.CalculatePercentages();

            int previous = calculatedSamples[0].TimeDifferenceInSeconds;
            DuplicatedTimes.Add(new Tuple<int, int>(previous, calculatedSamples[1].State));
            for (int i = 1; i < calculatedSamples.Count; i++)
            {
                DuplicatedTimes.Add(new Tuple<int, int>(calculatedSamples[i].TimeDifferenceInSeconds + previous, calculatedSamples[i].State));
                if (i < calculatedSamples.Count - 1)
                {
                    DuplicatedTimes.Add(new Tuple<int, int>(calculatedSamples[i].TimeDifferenceInSeconds + previous, calculatedSamples[i + 1].State));
                }
                previous = previous + calculatedSamples[i].TimeDifferenceInSeconds;
            }
        }
        public void CalculateHourlyStats(List<TimeStamp> calculatedSamples, int seconds, ExportOptions options)
        {
            int time = 0;
            int timeMark = 0;
            int indexCounter = 1;

            List<TimeStamp> hourSamples = new List<TimeStamp>();
            int[] indexes = new int[2];
            indexes[0] = indexCounter;
            foreach (TimeStamp sample in calculatedSamples)
            {
                time += sample.TimeDifferenceInSeconds;

                if (time > seconds) { throw new Exception("Invalid hour marks"); }

                hourSamples.Add(sample);

                if (time == seconds)
                {
                    indexes[1] = indexCounter;
                    timeMark++;
                    HourlyIndexes.Add(timeMark, indexes);
                    indexes = new int[2];
                    indexes[0] = indexCounter + 1;
                    Stats statHourly = new Stats();
                    statHourly.TotalTime = seconds;
                    foreach (int state in StatesMapping.Keys)
                    {
                        statHourly.StateTimes.Add(state, calculateStateTime(hourSamples, state));
                        statHourly.StateNumber.Add(state, calculateStateNumber(hourSamples, state));
                    }

                    // Calculate specific criteria states
                    foreach (KeyValuePair<int, int> entry in options.StateAndCriteria)
                    {
                        // OG
                        //statHourly.SpecificCrietriaStates.Add(entry.Key, calculateStateCriteriaNumber(hourSamples, entry.Key, entry.Value));

                        statHourly.SpecificTimeNumbers.Add(entry.Key, calculateStateCriteriaNumber(hourSamples, entry.Key, entry.Value));
                        statHourly.SpecificStateTimes.Add(entry.Key, calculateStateCriteriaTime(hourSamples, entry.Key, entry.Value));
                    }
                    foreach (KeyValuePair<int, int> entry in options.StateAndCriteriaAbove)
                    {
                        // OG
                        //statHourly.SpecificCrietriaStates.Add(entry.Key, calculateStateCriteriaNumber(hourSamples, entry.Key, entry.Value));

                        statHourly.SpecificStateNumbersAbove.Add(entry.Key, calculateStateCriteriaNumberAbove(hourSamples, entry.Key, entry.Value));
                        statHourly.SpecificStateTimesAbove.Add(entry.Key, calculateStateCriteriaTimeAbove(hourSamples, entry.Key, entry.Value));
                    }


                    statHourly.CalculatePercentages();
                    HourlyStats.Add(timeMark, statHourly);
                    time = 0;
                    hourSamples.Clear();
                }

                indexCounter++;
            }

            // Do last part (might be less than marked time)
            if (hourSamples.Count > 0)
            {
                timeMark++;
                indexes[1] = indexCounter - 1;
                HourlyIndexes.Add(timeMark, indexes);
                Stats statHourlyLast = new Stats();
                statHourlyLast.TotalTime = hourSamples.Sum(_sample => _sample.TimeDifferenceInSeconds);
                LastSectionTime = Math.Round((double)statHourlyLast.TotalTime / 60, 2);
                foreach (int state in StatesMapping.Keys)
                {
                    statHourlyLast.StateTimes.Add(state, calculateStateTime(hourSamples, state));
                    statHourlyLast.StateNumber.Add(state, calculateStateNumber(hourSamples, state));
                }

                // Calculate specific criteria states
                foreach (KeyValuePair<int, int> entry in options.StateAndCriteria)
                {
                    // OG
                    //statHourlyLast.SpecificCrietriaStates.Add(entry.Key, calculateStateCriteriaNumber(hourSamples, entry.Key, entry.Value));

                    statHourlyLast.SpecificTimeNumbers.Add(entry.Key, calculateStateCriteriaNumber(hourSamples, entry.Key, entry.Value));
                    statHourlyLast.SpecificStateTimes.Add(entry.Key, calculateStateCriteriaTime(hourSamples, entry.Key, entry.Value));
                }
                foreach (KeyValuePair<int, int> entry in options.StateAndCriteriaAbove)
                {
                    // OG
                    //statHourly.SpecificCrietriaStates.Add(entry.Key, calculateStateCriteriaNumber(hourSamples, entry.Key, entry.Value));

                    statHourlyLast.SpecificStateNumbersAbove.Add(entry.Key, calculateStateCriteriaNumberAbove(hourSamples, entry.Key, entry.Value));
                    statHourlyLast.SpecificStateTimesAbove.Add(entry.Key, calculateStateCriteriaTimeAbove(hourSamples, entry.Key, entry.Value));
                }

                statHourlyLast.CalculatePercentages();
                HourlyStats.Add(timeMark, statHourlyLast);
            }
        }
        public void ClearStats()
        {
            Stats = new Stats();
            HourlyStats = new Dictionary<int, Stats>();
            HourlyIndexes = new Dictionary<int, int[]>();
            DuplicatedTimes = new List<Tuple<int, int>>();
            StatesMapping = new Dictionary<int, string>();
        }

        // Private helpers
        private int calculateStateTime(List<TimeStamp> samples, int state)
        {
           return samples.Where((sample) => sample.State == state).Select((sample) => sample.TimeDifferenceInSeconds).Sum();
        }
        private int calculateStateNumber(List<TimeStamp> samples, int state, bool forTotal = false)
        {
            if (forTotal)
            {
                return samples.Count(sample => sample.State == state && !sample.IsMarker && !sample.IsTimeMarked);
            }
            return samples.Count(sample => sample.State == state);
        }
        private int calculateStateCriteriaTime(List<TimeStamp> samples, int state, int below)
        {
            return samples.Where((sample) => sample.State == state && sample.TimeDifferenceInSeconds <= below).Select((sample) => sample.TimeDifferenceInSeconds).Sum();
            
        }
        private int calculateStateCriteriaNumber(List<TimeStamp> samples, int state, int below)
        {
            return samples.Count(sample => sample.State == state && sample.TimeDifferenceInSeconds <= below);
        }
        private int calculateStateCriteriaTimeAbove(List<TimeStamp> samples, int state, int below)
        {
            return samples.Where((sample) => sample.State == state && sample.TimeDifferenceInSeconds >= below).Select((sample) => sample.TimeDifferenceInSeconds).Sum();

        }
        private int calculateStateCriteriaNumberAbove(List<TimeStamp> samples, int state, int below)
        {
            return samples.Count(sample => sample.State == state && sample.TimeDifferenceInSeconds >= below);
        }
    }
}
