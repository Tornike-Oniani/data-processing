using DataProcessing.Classes;
using DataProcessing.Classes.Calculate;
using DataProcessing.Constants;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
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
        private bool _isCalculateTotalSelected;
        private bool _setNameToClipboard;
        #endregion

        #region Private attributes
        private List<TimeStamp> records;
        private Func<Dictionary<string, int[]>> getDictionaryofFrequencyRanges;
        private Func<Dictionary<string, List<TimeStamp>>> getDataForAllSheets;
        #endregion

        #region Public properties
        public UserSelectedOptions options { get; set; }
        // All available timemarks for user to choose from combobox (10min, 20min, 30min, 1hr, 2hr, 4hr) we use function to convert string to seconds
        public List<string> TimeMarks { get; set; }
        // We use converter method (see below) to convert string from TimeMarks list into seconds
        public string SelectedTimeMark { get; set; }
        // Max number of states (can be 2 or 3 (in future we might also add 4))
        public List<string> RecordingTypes { get; set; }
        // We have to set recording type before calculating so the program knows how to calculate, recording types can be PS + Sleep + Wakefulness, Sleep + Wakefulness, Sleep + Wakefulness + Behaviors etc.
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
        public bool SetNameToClipboard
        {
            get { return _setNameToClipboard; }
            set { _setNameToClipboard = value; OnPropertyChanged("SetNameToClipboard"); }
        }
        // By what time margin should we define clusters (For example every time wakefulness is more than 10min)
        public int ClusterSeparationTime { get; set; }
        public bool IsCalculateTotalSelected
        {
            get { return _isCalculateTotalSelected; }
            set 
            { 
                _isCalculateTotalSelected = value; 
                // If this is selected the files get automatically saved based on SheetNumbers so we don't need to set name on clipboard and we wouldn't even know what to set since the filenames will be already decided
                if (_isCalculateTotalSelected)
                {
                    SetNameToClipboard = false;
                }
                OnPropertyChanged("IsCalculateTotalSelected"); 
            }
        }
        public ObservableCollection<ExportedFile> SheetFiles { get; set; }
        #endregion

        #region Commands
        public ICommand CalculateCommand { get; set; }
        #endregion

        #region Constructors
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

            SheetFiles = new ObservableCollection<ExportedFile>();
            string workfileName = WorkfileManager.GetInstance().SelectedWorkFile.Name;
            int sheetNumber = WorkfileManager.GetInstance().SelectedWorkFile.Sheets;
            for (int i = 0; i < sheetNumber; i++)
            {
                SheetFiles.Add(new ExportedFile() { SheetName = "Sheet" + (i + 1), FileName = workfileName + "_Sheet" + (i + 1) });
            }
            SheetFiles.Add(new ExportedFile() { SheetName = "Total", FileName = workfileName + "_Total" });


            // Set up commands
            CalculateCommand = new RelayCommand(Calculate);
        }
        #endregion

        #region Command actions
        public async void Calculate(object input = null)
        {
            if (!IsCalculateTotalSelected)
            {
                await CalculateSingleSheet();
                return;
            }

            // Select folder to export to
            string destination = Services.GetInstance().BrowserService.OpenFolderDialog();
            if (destination == null) { return; }

            await CalculateAllSheetsAndTotal(destination);
        }
        #endregion

        #region Public methods
        public void SetSelectedParams(List<TimeStamp> region, TimeSpan from, TimeSpan till, Func<Dictionary<string, int[]>> getDictionaryofFrequencyRanges, Func<Dictionary<string, List<TimeStamp>>> getDataForAllSheets)
        {
            this.records = region;
            this.From = from;
            this.Till = till;
            this.getDictionaryofFrequencyRanges = getDictionaryofFrequencyRanges;
            this.getDataForAllSheets = getDataForAllSheets;
        }
        #endregion

        #region Private helpers
        private async Task CalculateSingleSheet()
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

            // Export to excel
            CalculationOptions calcOptions = new CalculationOptions
                (
                    region,
                    new UserSelectedOptions()
                    {
                        SelectedTimeMark = SelectedTimeMark,
                        SelectedRecordingType = SelectedRecordingType,
                        ClusterSparationTime = ClusterSeparationTime,
                        FrequencyRanges = getDictionaryofFrequencyRanges(),
                        Criterias = criterias,
                        IsCalculateTotalSelected = IsCalculateTotalSelected
                    }
                );
            IDataProcessor dataProcessor;
            if (SelectedRecordingType == RecordingType.TwoStatesWithBehavior)
            {
                dataProcessor = new AnotherDataProcessor(calcOptions);
            }
            else
            {
                dataProcessor = new DataProcessor(calcOptions);
            }
            await new ExcelManager(
                calcOptions,
                dataProcessor.Calculate()
                ).ExportToExcelC();

            if (SetNameToClipboard)
            {
                Clipboard.SetText("Calc - " + WorkfileManager.GetInstance().SelectedWorkFile.Name);
            }
        }
        private async Task CalculateAllSheetsAndTotal(string destination)
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

            // Export to excel
            Dictionary<string, List<TimeStamp>> dataForAllSheets = getDataForAllSheets();

            foreach (KeyValuePair<string, List<TimeStamp>> entry in dataForAllSheets)
            {
                CalculationOptions calcOptions = new CalculationOptions
                (
                    entry.Value,
                    new UserSelectedOptions()
                    {
                        SelectedTimeMark = SelectedTimeMark,
                        SelectedRecordingType = SelectedRecordingType,
                        ClusterSparationTime = ClusterSeparationTime,
                        FrequencyRanges = getDictionaryofFrequencyRanges(),
                        Criterias = criterias,
                        IsCalculateTotalSelected = IsCalculateTotalSelected
                    }
                );
                IDataProcessor dataProcessor;
                if (SelectedRecordingType == RecordingType.TwoStatesWithBehavior)
                {
                    dataProcessor = new AnotherDataProcessor(calcOptions);
                }
                else
                {
                    dataProcessor = new DataProcessor(calcOptions);
                }
                string fileName = SheetFiles.FirstOrDefault(sf => sf.SheetName == entry.Key).FileName;
                await new ExcelManager(
                    calcOptions,
                    dataProcessor.Calculate()
                    ).ExportTotalSheetToExcelAndSave(destination, fileName);
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
        #endregion
    }
}
