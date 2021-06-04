using DataProcessing.Classes;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class ExportSettingsViewModel : BaseWindowViewModel
    {
        // Private attributes
        private List<TimeStamp> records;
        private bool _exportSelectedPeriod;

        // Public properties
        public List<float> TimeMarks { get; set; }
        public float SelectedTimeMark { get; set; }
        public int MaxStates { get; set; }
        public TimeSpan From { get; set; }
        public TimeSpan Till { get; set; }
        public int? WakefulnessBelow { get; set; }
        public int? ParadoxicalSleepBelow { get; set; }
        public bool ExportSelectedPeriod
        {
            get { return _exportSelectedPeriod; }
            set { _exportSelectedPeriod = value; OnPropertyChanged("ExportSelectedPeriod"); }
        }


        // Commands
        public ICommand ExportCommand { get; set; }
        public ICommand CancelCommand { get; set; }

        // Constructor
        public ExportSettingsViewModel(List<TimeStamp> records, TimeSpan from, TimeSpan till)
        {
            // Init
            this.Title = "Export settings";
            this.records = records;
            this.TimeMarks = new List<float>() { 0.5f, 1, 2 };
            this.SelectedTimeMark = TimeMarks[1];
            this.MaxStates = 3;
            this.From = from;
            this.Till = till;

            // Initialize commands
            ExportCommand = new RelayCommand(Export);
            CancelCommand = new RelayCommand(Cancel);
        }

        // Command actions
        public async void Export(object input = null)
        {
            ExportOptions exportOptions = new ExportOptions() { 
                TimeMark = SelectedTimeMark, MaxStates = MaxStates, 
                From = From, Till = Till
            };
            if (WakefulnessBelow != null) { exportOptions.StateAndCriteria.Add(MaxStates, (int)WakefulnessBelow); }
            if (ParadoxicalSleepBelow != null) { exportOptions.StateAndCriteria.Add(1, (int)ParadoxicalSleepBelow); }

            List<TimeStamp> markedRecords;
            if (ExportSelectedPeriod)
            {
                int fromCheck = records.Where(sample => sample.Time == From).ToList().Count;
                int tillCheck = records.Where(sample => sample.Time == Till).ToList().Count;
                if ( fromCheck == 0 || tillCheck == 0) { throw new Exception("Specified period doesn't exist!"); }
                markedRecords = records.Where(sample => isBetweenTimeInterval(From, Till, sample.Time)).ToList();
            }
            else
            {
                markedRecords = records;
            }
            markedRecords = AddTimeMarksToSamples(markedRecords);
            CalculateSamples(markedRecords);

            // 3. Calculate workfile intormation
            Workfile currentWorkfile = WorkfileManager.GetInstance().SelectedWorkFile;
            currentWorkfile.CalculateStats(markedRecords, exportOptions);
            currentWorkfile.CalculateHourlyStats(markedRecords, (int)TimeSpan.FromHours(SelectedTimeMark).TotalSeconds, exportOptions);

            // 4. Export to excel
            this.Window.Close();
            await new ExcelManager(exportOptions).ExportToExcel(markedRecords);
            currentWorkfile.ClearStats();
        }
        public void Cancel(object input = null)
        {
            this.Window.Close();
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
