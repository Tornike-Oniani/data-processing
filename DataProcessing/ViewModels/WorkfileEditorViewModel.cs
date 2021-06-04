using DataProcessing.Classes;
using DataProcessing.Models;
using DataProcessing.Repositories;
using DataProcessing.Utils;
using DataProcessing.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class WorkfileEditorViewModel : BaseViewModel
    {
        // Properties
        public DisplayManager DisplayManager { get; set; }
        public EntryManager EntryManager { get; set; }

        // Commands
        public ICommand OpenWorkfileDialogCommand { get; set; }
        public ICommand NewWorkfileDialogCommand { get; set; }
        public ICommand EditCommand { get; set; }
        public ICommand ExportCommand { get; set; }

        // Constructor
        public WorkfileEditorViewModel()
        {
            // Init
            DisplayManager = new DisplayManager();
            EntryManager = new EntryManager(DisplayManager.PopulateCommand);
            WorkfileManager.GetInstance().OnWorkfileChanged += SetupDisplayAndEntry;

            // Init commands
            OpenWorkfileDialogCommand = new RelayCommand(OpenWorkfileDialog);
            NewWorkfileDialogCommand = new RelayCommand(NewWorkfileDialog);
            EditCommand = new RelayCommand(Edit);
            ExportCommand = new RelayCommand(Export);
        }

        // Command actions
        public void OpenWorkfileDialog(object input = null)
        {
            Services.GetInstance().WindowService.OpenWindow(new OpenWorkfileViewModel());
        }
        public void NewWorkfileDialog(object input = null)
        {
            Services.GetInstance().WindowService.OpenWindow(new NewWorkfileViewModel());
        }
        public void Edit(object input)
        {
            if (input == null) return;
            string timeSpan = input as string;
            int[] times;

            if (!IsTimeSpanStringCorrect(timeSpan, out times)) { throw new Exception("Incorrect value. Correct format is hh:mm:ss"); }

            TimeSpan span = new TimeSpan(times[0], times[1], times[2]);
            DisplayManager.SelectedRow.Time = span;
            DisplayManager.SelectedRow.Update();
            DisplayManager.PopulateCommand.Execute(null);
        }
        public void Export(object input)
        {
            List<TimeStamp> samples = DisplayManager.Items.ToList();
            TimeSpan from = samples[0].Time;
            TimeSpan till = samples[samples.Count - 1].Time;
            if (DisplayManager.SelectedRows.Count > 1)
            {
                from = DisplayManager.SelectedRows[0].Time;
                till = DisplayManager.SelectedRows[DisplayManager.SelectedRows.Count - 1].Time;
            }
            Services.GetInstance().WindowService.OpenWindow(new ExportSettingsViewModel(DisplayManager.Items.ToList(), from, till));
        }

        // Event subscribers
        public void SetupDisplayAndEntry(Workfile workfile)
        {
            DisplayManager.PopulateCommand.Execute(null);
        }

        // Private helpers
        private bool IsTimeSpanStringCorrect(string timeSpan, out int[] times)
        {
            if (timeSpan.Length != 8) { times = null; return false; }
            string[] timeVals = timeSpan.Split(':');
            if (timeVals.Length != 3) { times = null; return false; }

            int[] numTimes = new int[3] { int.Parse(timeVals[0]), int.Parse(timeVals[1]), int.Parse(timeVals[2]) };
            times = numTimes;
            return true;
        }
    }
}
