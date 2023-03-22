using DataProcessing.Classes;
using DataProcessing.Classes.Calculate;
using DataProcessing.Constants;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    internal class OptionsViewModel : BaseViewModel
    {
        #region Property fields
        private bool _exportSelectedPeriod;
        private TimeSpan _from;
        private TimeSpan _till;
        private string _selectedRecordingType;
        #endregion

        #region Private attributes
        private List<TimeStamp> records;
        private Func<Dictionary<string, int[]>> getDictionaryofFrequencyRanges;
        #endregion

        #region Public properties
        public UserSelectedOptions options { get; set; }
        // All available timemarks for user to choose from combobox (10min, 20min, 30min, 1hr, 2hr, 4hr) we use function to convert string to seconds
        public List<string> TimeMarks { get; set; }
        // We use converter method (see below) to convert string from TimeMarks list into seconds
        public string SelectedTimeMark { get; set; }
        // Max number of states (can be 2 or 3 (in future we might also add 4))
        public List<string> RecordingTypes { get; set; }
        public string SelectedRecordingType
        {
            get { return _selectedRecordingType; }
            set
            {
                _selectedRecordingType = value;
                // If user has set these criterias and then changed state to 2
                // where these criterias don't exist then set them to null so they 
                // won't be calculated (We also set Visibility to collapsed on UI
                // with ValueConverter)
                if (value == RecordingType.TwoStates)
                {
                    ParadoxicalSleepBelow = null;
                    ParadoxicalSleepAbove = null;
                }
                OnPropertyChanged("SelectedRecordingType");
            }
        }
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
        #endregion

        #region Commands
        public ICommand CalculateCommand { get; set; }
        #endregion

        // Constructor
        public OptionsViewModel()
        {
            // Init
            TimeMarks = new List<string>() { "10min", "15min", "20min", "30min", "1hr", "2hr", "4hr" };
            SelectedTimeMark = TimeMarks[3];
            RecordingTypes = new List<string>()
            {
                RecordingType.ThreeStates,
                RecordingType.TwoStates,
                RecordingType.TwoStatesWithBehavior
            };
            SelectedRecordingType = RecordingTypes[0];

            // Set up commands
            CalculateCommand = new RelayCommand(Calculate);
        }

        #region Command actions
        public async void Calculate(object input = null)
        {
            // Set specific criterias based on max states
            List<SpecificCriteria> criterias = new List<SpecificCriteria>();
            if (SelectedRecordingType == RecordingType.ThreeStates)
            {
                criterias = new List<SpecificCriteria>()
                {
                    new SpecificCriteria() { State = 3, Operand = "Below", Value = WakefulnessBelow },
                    new SpecificCriteria() { State = 2, Operand = "Below", Value = SleepBelow },
                    new SpecificCriteria() { State = 1, Operand = "Below", Value = ParadoxicalSleepBelow },
                    new SpecificCriteria() { State = 3, Operand = "Above", Value = WakefulnessAbove },
                    new SpecificCriteria() { State = 2, Operand = "Above", Value = SleepAbove },
                    new SpecificCriteria() { State = 1, Operand = "Above", Value = ParadoxicalSleepAbove },
                };
            }
            else if (SelectedRecordingType == RecordingType.TwoStates)
            {
                criterias = new List<SpecificCriteria>()
                {
                    new SpecificCriteria() { State = 2, Operand = "Below", Value = WakefulnessBelow },
                    new SpecificCriteria() { State = 1, Operand = "Below", Value = SleepBelow },
                    new SpecificCriteria() { State = 2, Operand = "Above", Value = WakefulnessAbove },
                    new SpecificCriteria() { State = 1, Operand = "Above", Value = SleepAbove },
                };
            }

            ExcelResources.GetInstance().MaxStates = RecordingType.MaxStates[SelectedRecordingType];

            List<TimeStamp> region;
            if (ExportSelectedPeriod)
            {
                int fromCheck = records.Where(sample => sample.Time == From).ToList().Count;
                int tillCheck = records.Where(sample => sample.Time == Till).ToList().Count;
                if (fromCheck == 0 || tillCheck == 0) { throw new Exception("Specified period doesn't exist!"); }
                region = records.Where(sample => isBetweenTimeInterval(From, Till, sample.Time)).ToList();
            }
            else
            {
                region = records;
            }

            // 4. Export to excel
            CalculationOptions calcOptions = new CalculationOptions
                (
                region,
                new UserSelectedOptions()
                {
                    SelectedTimeMark = SelectedTimeMark,
                    SelectedRecordingType = SelectedRecordingType,
                    ClusterSparationTime = ClusterSeparationTime,
                    FrequencyRanges = getDictionaryofFrequencyRanges(),
                    Criterias = criterias
                });
            DataProcessor dataProcessor = new DataProcessor(calcOptions);
            await new ExcelManager(
                calcOptions,
                dataProcessor.Calculate()
                ).ExportToExcelC();

            if (SetNameToClipboard)
            {
                Clipboard.SetText("Calc - " + WorkfileManager.GetInstance().SelectedWorkFile.Name);
            }
        }
        #endregion

        #region Public methods
        public void SetSelectedParams(List<TimeStamp> region, TimeSpan from, TimeSpan till, Func<Dictionary<string, int[]>> getDictionaryofFrequencyRanges)
        {
            this.records = region;
            this.From = from;
            this.Till = till;
            this.getDictionaryofFrequencyRanges = getDictionaryofFrequencyRanges;
        }
        #endregion

        #region Private helpers
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
        #endregion
    }
}
