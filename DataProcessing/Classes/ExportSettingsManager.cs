using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace DataProcessing.Classes
{
    class ExportSettingsManager : ObservableObject
    {
        // Private attributes
        private List<TimeStamp> records;
        private bool _exportSelectedPeriod;
        private Dictionary<string, int[]> customFrequencyRanges;
        private TimeSpan _from;
        private TimeSpan _till;

        // Public properties
        public List<float> TimeMarks { get; set; }
        public float SelectedTimeMark { get; set; }
        public int MaxStates { get; set; }
        public TimeSpan From
        {
            get { return _from; }
            set { _from = value; OnPropertyChanged("From"); }
        }     
        public TimeSpan Till
        {
            get { return _till; }
            set { _till = value; OnPropertyChanged("Till"); }
        }
        public int? WakefulnessBelow { get; set; }
        public int? SleepBelow { get; set; }
        public int? ParadoxicalSleepBelow { get; set; }
        public int? WakefulnessAbove { get; set; }
        public int? SleepAbove { get; set; }
        public int? ParadoxicalSleepAbove { get; set; }
        public bool ExportSelectedPeriod
        {
            get { return _exportSelectedPeriod; }
            set { _exportSelectedPeriod = value; OnPropertyChanged("ExportSelectedPeriod"); }
        }
        public bool SetNameToClipboard { get; set; }


        // Commands
        public ICommand ExportCommand { get; set; }

        public ExportSettingsManager(

            )
        {
            this.TimeMarks = new List<float>() { 0.5f, 1, 2, 4 };
            this.SelectedTimeMark = TimeMarks[1];
            this.MaxStates = 3;

            ExportCommand = new RelayCommand(Export);
        }

        public async void Export(object input = null)
        {
            ExportOptions exportOptions = new ExportOptions()
            {
                TimeMark = SelectedTimeMark,
                MaxStates = MaxStates,
                From = From,
                Till = Till,
                Criterias = new List<SpecificCriteria>()
                {
                    new SpecificCriteria() { State = MaxStates, Operand = "Below", Value = WakefulnessBelow },
                    new SpecificCriteria() { State = 2, Operand = "Below", Value = SleepBelow },
                    new SpecificCriteria() { State = 1, Operand = "Below", Value = ParadoxicalSleepBelow },
                    new SpecificCriteria() { State = MaxStates, Operand = "Above", Value = WakefulnessAbove },
                    new SpecificCriteria() { State = 2, Operand = "Above", Value = SleepAbove },
                    new SpecificCriteria() { State = 1, Operand = "Above", Value = ParadoxicalSleepAbove },
                },
                customFrequencyRanges = customFrequencyRanges
            };

            List<TimeStamp> markedRecords;
            if (ExportSelectedPeriod)
            {
                int fromCheck = records.Where(sample => sample.Time == From).ToList().Count;
                int tillCheck = records.Where(sample => sample.Time == Till).ToList().Count;
                if (fromCheck == 0 || tillCheck == 0) { throw new Exception("Specified period doesn't exist!"); }
                markedRecords = records.Where(sample => isBetweenTimeInterval(From, Till, sample.Time)).ToList();
            }
            else
            {
                markedRecords = records;
            }
            // Keep non marked original records for total frequency calculation in DataProcessor
            List<TimeStamp> nonMarkedRecords = new List<TimeStamp>();
            foreach (TimeStamp timeStamp in markedRecords)
            {
                nonMarkedRecords.Add(timeStamp.Clone());
            }
            markedRecords = AddTimeMarksToSamples(markedRecords);
            CalculateSamples(markedRecords);
            CalculateSamples(nonMarkedRecords);

            // 4. Export to excel
            DataProcessor dataProcessor = new DataProcessor(markedRecords, exportOptions);
            // We are also passing non marked records for total frequencies
            dataProcessor.Calculate(nonMarkedRecords);
            await new ExcelManager(exportOptions,
                dataProcessor.CreateStatTables(),
                dataProcessor.CreateGraphTables(),
                dataProcessor.CreateFrequencyTables(),
                dataProcessor.CreateLatencyTable(),
                dataProcessor.CreateCustomFrequencyTables()).
                ExportToExcel(
                    markedRecords,
                    dataProcessor.getDuplicatedTimes(),
                    dataProcessor.getHourRowIndexes(),
                    dataProcessor.getHourRowIndexesTime());
            if (SetNameToClipboard)
                Clipboard.SetText("Calc - " + WorkfileManager.GetInstance().SelectedWorkFile.Name);
        }

        public void SetSettings(
            List<TimeStamp> records,
            TimeSpan from,
            TimeSpan till,
            Dictionary<string, int[]> customFrequencyRanges)
        {
            this.records = records;
            this.From = from;
            this.Till = till;
            this.customFrequencyRanges = customFrequencyRanges;
        }
        // Private helpers
        private List<TimeStamp> AddTimeMarksToSamples(List<TimeStamp> records)
        {
            List<TimeStamp> result = new List<TimeStamp>();

            TimeSpan markCap = TimeSpan.FromHours(SelectedTimeMark);
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

                if (span == new TimeSpan(23, 50, 16))
                {
                    Console.WriteLine("Break point");
                }

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
                        result.Add(new TimeStamp() { Time = timeMark, State = state, IsTimeMarked = true });
                        lastMarkTime = timeMark;
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
        private bool isBetweenTimeInterval(TimeSpan from, TimeSpan till, TimeSpan time)
        {
            if (from < till)
            {
                return from <= time && time <= till;
            }
            else
            {
                return from <= time || time <= till;
            }
        }
    }
}
