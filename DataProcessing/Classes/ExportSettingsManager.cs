using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
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
        // All available timemarks for user to choose from combobox (10min, 20min, 1hr, 2hr, 4hr)
        public List<string> TimeMarks { get; set; }
        // We use converter method (see below) to convert string from TimeMarks list into seconds
        public string SelectedTimeMark { get; set; }
        // Max number of states (can be 2 or 3 (in future we might also add 4))
        public List<int> States { get; set; }
        public int SelectedState { get; set; }
        // Selected period from data to process
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
        // Specific crieterias for stat calculations
        public int? WakefulnessBelow { get; set; }
        public int? SleepBelow { get; set; }
        public int? ParadoxicalSleepBelow { get; set; }
        public int? WakefulnessAbove { get; set; }
        public int? SleepAbove { get; set; }
        public int? ParadoxicalSleepAbove { get; set; }
        // Check if user wishes to process and export whole data or a selected period
        public bool ExportSelectedPeriod
        {
            get { return _exportSelectedPeriod; }
            set { _exportSelectedPeriod = value; OnPropertyChanged("ExportSelectedPeriod"); }
        }
        // Check if user wishes to set filename on clipboard (We need this because file name can't be set on opened excel file by interop)
        public bool SetNameToClipboard { get; set; }
        // By what time margin should we define clusters (For example every time wakefulness is more than 10min)
        public int ClusterSeparationTime { get; set; }

        // Commands
        public ICommand ExportCommand { get; set; }

        public ExportSettingsManager()
        {
            // Init
            // 10min, 20min, 1hr, 2hr and 4hr in seconds
            this.TimeMarks = new List<string>() { "10min", "20min", "30min", "1hr", "2hr", "4hr" };
            this.SelectedTimeMark = TimeMarks[3];
            this.States = new List<int>() { 2, 3 };
            this.SelectedState = States[1];

            // Set up commands
            ExportCommand = new RelayCommand(Export);
        }

        public async void Export(object input = null)
        {
            List<SpecificCriteria> criterias = new List<SpecificCriteria>();
            if (SelectedState == 3)
            {
                criterias = new List<SpecificCriteria>()
                {
                    new SpecificCriteria() { State = SelectedState, Operand = "Below", Value = WakefulnessBelow },
                    new SpecificCriteria() { State = 2, Operand = "Below", Value = SleepBelow },
                    new SpecificCriteria() { State = 1, Operand = "Below", Value = ParadoxicalSleepBelow },
                    new SpecificCriteria() { State = SelectedState, Operand = "Above", Value = WakefulnessAbove },
                    new SpecificCriteria() { State = 2, Operand = "Above", Value = SleepAbove },
                    new SpecificCriteria() { State = 1, Operand = "Above", Value = ParadoxicalSleepAbove },
                };
            }
            else if (SelectedState == 2)
            {
                criterias = new List<SpecificCriteria>()
                {
                    new SpecificCriteria() { State = SelectedState, Operand = "Below", Value = WakefulnessBelow },
                    new SpecificCriteria() { State = 1, Operand = "Below", Value = SleepBelow },
                    new SpecificCriteria() { State = SelectedState, Operand = "Above", Value = WakefulnessAbove },
                    new SpecificCriteria() { State = 1, Operand = "Above", Value = SleepAbove },
                };
            }
            ExportOptions exportOptions = new ExportOptions()
            {
                // Init
                TimeMark = ConvertTimeMarkToSeconds(SelectedTimeMark),
                TimeMarkInSeconds = ConvertTimeMarkToSeconds(SelectedTimeMark),
                MaxStates = SelectedState,
                From = From,
                Till = Till,
                // Set up criterias for processing
                Criterias = criterias,
                customFrequencyRanges = customFrequencyRanges,
                ClusterSeparationTimeInSeconds = ClusterSeparationTime * 60
            };

            ExcelResources.GetInstance().MaxStates = SelectedState;

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
            DataProcessor dataProcessor = new DataProcessor(markedRecords, nonMarkedRecords, exportOptions);
            await new ExcelManager(
                exportOptions,
                dataProcessor.Calculate(),
                markedRecords,
                nonMarkedRecords
                ).ExportToExcelC();

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

            //TimeSpan markCap1 = TimeSpan.FromHours(1);
            TimeSpan markCap = TimeSpan.FromSeconds(ConvertTimeMarkToSeconds(SelectedTimeMark));
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
        private int ConvertTimeMarkToSeconds(string timeMark)
        {
            switch (timeMark)
            {
                case "10min":
                    return 600;
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
    }
}
