﻿using DataProcessing.Classes;
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
        // Private attributes
        private bool _customRangesEnabled;
        private string _frequencyRange;
        private int _selectedFrequencyTimeUnit;
        private bool _isRangeEntryFocused;
        private FrequencyRangeTemplateManager frequencyRangeTemplateManager;
        private FrequencyRangeTemplate _selectedFrequencyRangeTemplate;
        private bool _isTemplateSelected;
        private bool _isTemplateChanged;
        private int _selectedTabIndex;

        // Properties
        public DisplayManager DisplayManager { get; set; }
        public ExportSettingsManager ExportSettingsManager { get; set; }
        public bool CustomRangesEnabled
        {
            get { return _customRangesEnabled; }
            set { _customRangesEnabled = value; OnPropertyChanged("CustomRangesEnabled"); IsRangeEntryFocused = value; }
        }
        public string FrequencyRange
        {
            get { return _frequencyRange; }
            set { _frequencyRange = value; OnPropertyChanged("FrequencyRange"); }
        }
        public List<int> FrequencyTimeUnits { get; set; }
        public int SelectedFrequencyTimeUnit
        {
            get { return _selectedFrequencyTimeUnit; }
            set { _selectedFrequencyTimeUnit = value; OnPropertyChanged("SelectedFrequencyTimeUnit"); }
        }
        public ObservableCollection<FrequencyRange> FrequencyRanges { get; set; }
        public FrequencyRange SelectedFrequencyRange { get; set; }
        public bool IsRangeEntryFocused
        {
            get { return _isRangeEntryFocused; }
            set { _isRangeEntryFocused = value; OnPropertyChanged("IsRangeEntryFocused"); }
        }
        public ObservableCollection<FrequencyRangeTemplate> FrequencyRangeTemplates { get; set; }
        public FrequencyRangeTemplate SelectedFrequencyRangeTemplate
        {
            get { return _selectedFrequencyRangeTemplate; }
            set
            {
                _selectedFrequencyRangeTemplate = value;
                OnPropertyChanged("SelectedFrequencyRangeTemplate");
                FrequencyRanges.Clear();
                IsTemplateChanged = false;
                if (_selectedFrequencyRangeTemplate == null) { IsTemplateSelected = false; return; }
                IsTemplateSelected = true;
                if (_selectedFrequencyRangeTemplate.FrequencyRanges == null) { return; }
                foreach (FrequencyRange range in _selectedFrequencyRangeTemplate.FrequencyRanges)
                {
                    FrequencyRanges.Add(range);
                }
            }
        }
        public bool IsTemplateSelected
        {
            get { return _isTemplateSelected; }
            set { _isTemplateSelected = value; OnPropertyChanged("IsTemplateSelected"); }
        }
        public bool IsTemplateChanged
        {
            get { return _isTemplateChanged; }
            set { _isTemplateChanged = value; OnPropertyChanged("IsTemplateChanged"); }
        }
        public int SelectedTabIndex
        {
            get { return _selectedTabIndex; }
            set
            {
                _selectedTabIndex = value;
                OnPropertyChanged("SelectedTabIndex");
                if (value == 2)
                    SetExportSettings();
            }
        }

        // Commands
        public ICommand AddRangeCommand { get; set; }
        public ICommand RemoveRangeCommand { get; set; }
        public ICommand NewTemplateCommand { get; set; }
        public ICommand DeleteTemplateCommand { get; set; }
        public ICommand SaveTemplateCommand { get; set; }
        public ICommand SaveTemplateAsCommand { get; set; }
        public ICommand NextCommand { get; set; }
        public ICommand PrevCommand { get; set; }

        // Constructor
        public WorkfileEditorViewModel()
        {
            // Init
            DisplayManager = new DisplayManager();
            WorkfileManager.GetInstance().OnWorkfileChanged += SetupDisplayAndEntry;
            ExportSettingsManager = new ExportSettingsManager();
            // These values gets converted into 'sec', 'min' and 'hr' with converter
            FrequencyTimeUnits = new List<int>() { 1, 60, 3600 };
            SelectedFrequencyTimeUnit = FrequencyTimeUnits[0];
            FrequencyRanges = new ObservableCollection<FrequencyRange>();
            FrequencyRangeTemplates = new ObservableCollection<FrequencyRangeTemplate>();
            IsRangeEntryFocused = false;
            frequencyRangeTemplateManager = new FrequencyRangeTemplateManager();
            LoadFrequencyRangeTempaltes();

            // Init commands
            AddRangeCommand = new RelayCommand(AddRange);
            RemoveRangeCommand = new RelayCommand(RemoveRange);
            SaveTemplateCommand = new RelayCommand(SaveTemplate);
            SaveTemplateAsCommand = new RelayCommand(SaveTemplateAs);
            NewTemplateCommand = new RelayCommand(NewTemplate);
            DeleteTemplateCommand = new RelayCommand(DeleteTemplate);
            NextCommand = new RelayCommand(Next, CanNext);
            PrevCommand = new RelayCommand(Prev, CanPrev);
        }

        // Command actions
        public void AddRange(object input = null)
        {
            string[] rangeSplit = FrequencyRange.Trim().Split('-');
            // Check for incorrect range entry
            if (
                FrequencyRange.Trim().Any(Char.IsWhiteSpace) ||
                rangeSplit.Length != 2 ||
                !int.TryParse(rangeSplit[0], out _) ||
                !int.TryParse(rangeSplit[1], out _)) { return; }

            FrequencyRange range = new FrequencyRange() { Range = FrequencyRange, TimeUnit = SelectedFrequencyTimeUnit };

            // Check for duplicate entry
            if (FrequencyRanges.Any(r => r.Range == FrequencyRange && r.TimeUnit == SelectedFrequencyTimeUnit))
            {
                FrequencyRange = null;
                IsRangeEntryFocused = false;
                IsRangeEntryFocused = true;
                return;
            }

            FrequencyRanges.Add(range);
            FrequencyRange = null;
            IsRangeEntryFocused = false;
            IsRangeEntryFocused = true;
            IsTemplateChanged = true;
        }
        public void RemoveRange(object input = null)
        {
            if (SelectedFrequencyRange == null) { return; }
            FrequencyRanges.Remove(SelectedFrequencyRange);
            IsTemplateChanged = true;
        }
        public void SaveTemplate(object input = null)
        {
            SelectedFrequencyRangeTemplate.FrequencyRanges = FrequencyRanges.ToList();
            frequencyRangeTemplateManager.SaveFrequencyRangeTemplates(FrequencyRangeTemplates.ToList());
            IsTemplateChanged = false;
        }
        public void SaveTemplateAs(object input = null)
        {
            string templateName = Services.GetInstance().DialogService.OpenTextDialog("Template name:");
            if (templateName == null) { return; }
            if (FrequencyRangeTemplates.Any(t => t.Name == templateName))
            {
                MessageBox.Show("Template with that name already exists.", "Save as", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            FrequencyRangeTemplate newTemplate = new FrequencyRangeTemplate() { Name = templateName };
            newTemplate.FrequencyRanges = FrequencyRanges.ToList();
            FrequencyRangeTemplates.Add(newTemplate);
            SelectedFrequencyRangeTemplate = newTemplate;
            frequencyRangeTemplateManager.SaveFrequencyRangeTemplates(FrequencyRangeTemplates.ToList());
            IsTemplateChanged = false;
        }
        public void NewTemplate(object input = null)
        {
            string templateName = Services.GetInstance().DialogService.OpenTextDialog("Template name:");
            if (templateName == null) { return; }
            if (FrequencyRangeTemplates.Any(t => t.Name == templateName))
            {
                MessageBox.Show("Template with that name already exists.", "Save as", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            FrequencyRangeTemplate newTemplate = new FrequencyRangeTemplate() { Name = templateName };
            FrequencyRangeTemplates.Add(newTemplate);
            SelectedFrequencyRangeTemplate = newTemplate;
            frequencyRangeTemplateManager.SaveFrequencyRangeTemplates(FrequencyRangeTemplates.ToList());
            IsTemplateChanged = false;
        }
        public void DeleteTemplate(object input = null)
        {
            if (MessageBox.Show($"Are you sure you want to delete '{SelectedFrequencyRangeTemplate.Name}'?", "Delete template", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                FrequencyRangeTemplates.Remove(SelectedFrequencyRangeTemplate);
                frequencyRangeTemplateManager.SaveFrequencyRangeTemplates(FrequencyRangeTemplates.ToList());
                if (FrequencyRangeTemplates.Count == 0) { SelectedFrequencyRangeTemplate = null; return; }
                SelectedFrequencyRangeTemplate = FrequencyRangeTemplates[0];
                IsTemplateChanged = false;
            }
        }
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

        // Event subscribers
        public void SetupDisplayAndEntry(Workfile workfile)
        {
            DisplayManager.PopulateCommand.Execute(null);
        }

        // Private helpers
        private Dictionary<string, int[]> FrequencyRangesToArray()
        {
            // If adding custom ranges is disabled or not template is selected return blank list
            if (!CustomRangesEnabled || !IsTemplateSelected || FrequencyRanges.Count == 0) { return new Dictionary<string, int[]>(); }

            Dictionary<string, int[]> result = new Dictionary<string, int[]>();
            List<FrequencyRange> ranges = FrequencyRanges.ToList();
            ranges = ranges.OrderBy(r => int.Parse(r.Range.Split('-')[0])).ToList();
            string[] rangeSplit;
            foreach (FrequencyRange frequencyRange in ranges)
            {
                rangeSplit = frequencyRange.Range.Split('-');
                // Convert range into seconds
                result.Add(frequencyRange.Range,
                    new int[2]
                    {
                        int.Parse(rangeSplit[0]) * SelectedFrequencyTimeUnit,
                        int.Parse(rangeSplit[1]) * SelectedFrequencyTimeUnit
                    });
            }
            rangeSplit = ranges.Last().Range.Split('-');
            // Add more than last interval (if last interval is 15-20 we add >20)
            result.Add($">{rangeSplit[1]}",
                new int[2]
                {
                    int.Parse(rangeSplit[1]) * SelectedFrequencyTimeUnit,
                    int.MaxValue
                });
            return result;
        }
        private void LoadFrequencyRangeTempaltes()
        {
            foreach (FrequencyRangeTemplate frequencyRangeTemplate in frequencyRangeTemplateManager.GetFrequencyRangeTemplates())
            {
                FrequencyRangeTemplates.Add(frequencyRangeTemplate);
            }
            if (FrequencyRangeTemplates.Count != 0) { SelectedFrequencyRangeTemplate = FrequencyRangeTemplates[0]; }
        }
        private void SetExportSettings()
        {
            List<TimeStamp> samples = DisplayManager.Items.ToList();
            TimeSpan from = samples[0].Time;
            TimeSpan till = samples[samples.Count - 1].Time;
            if (DisplayManager.SelectedRows.Count > 1)
            {
                from = DisplayManager.SelectedRows[0].Time;
                till = DisplayManager.SelectedRows[DisplayManager.SelectedRows.Count - 1].Time;
            }
            ExportSettingsManager.SetSettings(DisplayManager.Items.ToList(), from, till, FrequencyRangesToArray());
        }
    }
}
