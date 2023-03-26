using DataProcessing.Classes;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class WorkfileEditorViewModel : BaseViewModel
    {
        #region Property fields
        private int _selectedTabIndex;
        #endregion

        #region Properties
        public DisplayManager DisplayManager { get; set; }
        public OptionsViewModel OptionsViewModel { get; set; }
        public FrequencyRangesViewModel FrequencyRangesViewModel { get; set; }
        public int SelectedTabIndex
        {
            get { return _selectedTabIndex; }
            set
            {
                _selectedTabIndex = value;
                OnPropertyChanged("SelectedTabIndex");
                if (value == 2)
                {
                    SetSelectedParams();
                }
            }
        }
        #endregion

        #region Commands
        public ICommand NextCommand { get; set; }
        public ICommand PrevCommand { get; set; }
        #endregion

        #region Constructors
        public WorkfileEditorViewModel()
        {
            // Init
            DisplayManager = new DisplayManager();
            WorkfileManager.GetInstance().OnWorkfileChanged += SetupDisplayAndEntry;
            FrequencyRangesViewModel = new FrequencyRangesViewModel();
            OptionsViewModel = new OptionsViewModel();

            // Init commands
            NextCommand = new RelayCommand(Next, CanNext);
            PrevCommand = new RelayCommand(Prev, CanPrev);
        }
        #endregion

        #region Command actions
        public void Next(object input = null)
        {
            SelectedTabIndex += 1;
        }
        public bool CanNext(object input = null)
        {
            return SelectedTabIndex < 2;
        }
        public void Prev(object input = null)
        {
            SelectedTabIndex -= 1;
        }
        public bool CanPrev(object input = null)
        {
            return SelectedTabIndex > 0;
        }
        #endregion

        #region Event subscribers
        public void SetupDisplayAndEntry(Workfile workfile)
        {
            DisplayManager.PopulateCommand.Execute(null);
        }
        #endregion

        #region Private helpers
        private void SetSelectedParams()
        {
            List<TimeStamp> samples = DisplayManager.Items.ToList();
            TimeSpan from = samples[0].Time;
            TimeSpan till = samples[samples.Count - 1].Time;
            if (DisplayManager.SelectedRows.Count > 1)
            {
                from = DisplayManager.SelectedRows[0].Time;
                till = DisplayManager.SelectedRows[DisplayManager.SelectedRows.Count - 1].Time;
            }
            //ExportSettingsManager.SetSettings(DisplayManager.Items.ToList(), from, till, FrequencyRangesViewModel.FrequencyRangesToArray());
            OptionsViewModel.SetSelectedParams(samples, from, till, FrequencyRangesViewModel.FrequencyRangesToArray);

        }
        #endregion
    }
}
