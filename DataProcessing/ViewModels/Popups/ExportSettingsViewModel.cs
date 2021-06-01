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
        // Public properties
        public List<float> TimeMarks { get; set; }
        public float SelectedTimeMark { get; set; }
        public int MaxStates { get; set; }
        public TimeSpan DarkFrom { get; set; }
        public TimeSpan DarkTill { get; set; }

        // Commands
        public ICommand ExportCommand { get; set; }
        public ICommand CancelCommand { get; set; }

        // Constructor
        public ExportSettingsViewModel()
        {
            // Init
            this.TimeMarks = new List<float>() { 0.5f, 1, 2 };
            this.SelectedTimeMark = TimeMarks[1];
            this.MaxStates = 4;
            this.DarkFrom = new TimeSpan(10, 0, 0);
            this.DarkTill = new TimeSpan(22, 0, 0);

            // Initialize commands
            ExportCommand = new RelayCommand(Export);
            CancelCommand = new RelayCommand(Cancel);
        }

        // Command actions
        public void Export(object input = null)
        {

        }
        public void Cancel(object input = null)
        {

        }
    }
}
