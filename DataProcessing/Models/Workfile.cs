using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Models
{
    class Workfile
    {
        private List<int> states = new List<int>();

        public int Id { get; set; }
        public string Name { get; set; }
        public Stats Stats { get; set; } = new Stats();
        public Dictionary<int, Stats> HourlyStats { get; set; } = new Dictionary<int, Stats>();
        public Dictionary<int, int[]> HourlyIndexes { get; set; } = new Dictionary<int, int[]>();
        public List<Tuple<int, int>> DuplicatedTimes { get; set; } = new List<Tuple<int, int>>();
        public Dictionary<int, string> StatesMapping { get; set; }

        public void CalculateStats()
        {
            List<DataSample> samples = DataSample.Find();

            // Calculate total stats
            Stats.TotalTime = samples.Sum((sample) => sample.D);
            states = samples.Where(sample => sample.State != 0).Select(sample => sample.State).Distinct().ToList();
            states.Sort();
            foreach (int state in states)
            {
                Stats.StateTimes.Add(state, calculateStateTime(samples, state));
            }
            Stats.CalculatePercentages();

            samples = DataSample.Find();

            int previous = samples[0].D;
            DuplicatedTimes.Add(new Tuple<int, int>(previous, 1));
            for (int i = 1; i < samples.Count; i++)
            {
                DuplicatedTimes.Add(new Tuple<int, int>(samples[i].D + previous, samples[i].State));
                if (i < samples.Count - 1)
                {
                    DuplicatedTimes.Add(new Tuple<int, int>(samples[i].D + previous, samples[i + 1].State));
                }
                previous = previous + samples[i].D;
            }

            // Map states and phases
            if (states.Count == 3)
            {
                StatesMapping = new Dictionary<int, string>();
                StatesMapping.Add(1, "Paradoxical sleep");
                StatesMapping.Add(2, "Sleep"); // ???
                StatesMapping.Add(3, "Wakefulness");

            }
            else if (states.Count == 4)
            {
                StatesMapping = new Dictionary<int, string>();
                StatesMapping.Add(1, "Paradoxical sleep");
                StatesMapping.Add(2, "Deep sleep");
                StatesMapping.Add(3, "Light sleep");
                StatesMapping.Add(4, "Wakefulness");
            }
            else
            {
                throw new Exception($"File has {states.Count} states");
            }
        }
        public void CalculateHourlyStats()
        {
            List<DataSample> samples = DataSample.Find();

            int time = 0;
            int timeMark = 0;
            int indexCounter = 1;

            List<DataSample> hourSamples = new List<DataSample>();
            int[] indexes = new int[2];
            indexes[0] = indexCounter;
            foreach (DataSample sample in samples)
            {
                time += sample.D;
                if (time > 3600) { throw new Exception("Invalid 1 hour marks"); }

                hourSamples.Add(sample);

                if (time == 3600)
                {
                    indexes[1] = indexCounter;
                    timeMark++;
                    HourlyIndexes.Add(timeMark, indexes);
                    indexes = new int[2];
                    indexes[0] = indexCounter + 1;
                    Stats statHourly = new Stats();
                    statHourly.TotalTime = 3600;
                    foreach (int state in states)
                    {
                        statHourly.StateTimes.Add(state, calculateStateTime(hourSamples, state));
                    }
                    statHourly.CalculatePercentages();
                    HourlyStats.Add(timeMark, statHourly);
                    time = 0;
                    hourSamples.Clear();
                }

                indexCounter++;
            }

            timeMark++;
            indexes[1] = indexCounter;
            HourlyIndexes.Add(timeMark, indexes);
            Stats statHourlyLast = new Stats();
            statHourlyLast.TotalTime = hourSamples.Sum(_sample => _sample.D);
            foreach (int state in states)
            {
                statHourlyLast.StateTimes.Add(state, calculateStateTime(hourSamples, state));
            }
            statHourlyLast.CalculatePercentages();
            HourlyStats.Add(timeMark, statHourlyLast);
        }

        private int calculateStateTime(List<DataSample> samples, int state)
        {
           return samples.Where((sample) => sample.State == state).Select((sample) => sample.D).Sum();
        }
    }
}
