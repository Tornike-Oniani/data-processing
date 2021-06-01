using DataProcessing.Utils;
using DataProcessing.Utils.Interfaces;
using DataProcessing.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class RadioDialogViewModel : BaseViewModel
    {
        private IWindow window;

        // Public properties
        public bool Is3Checked { get; set; }
        public bool Is4Checked { get; set; }
        public int DialogResult { get; set; }

        // Commands
        public ICommand OkCommand { get; set; }
        public ICommand CancelCommand { get; set; }

        // Constructor
        public RadioDialogViewModel(IWindow window)
        {
            // Init
            this.window = window;

            // Initialize commands
            OkCommand = new RelayCommand(Ok);
            CancelCommand = new RelayCommand(Cancel);
        }

        // Command actions
        public void Ok(object input = null)
        {
            if (!Is3Checked && !Is4Checked) { throw new Exception("Please select the state mode"); }
            if (Is3Checked) { DialogResult = 3; }
            if (Is4Checked) { DialogResult = 4; }
            window.Close();
        }
        public void Cancel(object input = null)
        {
            DialogResult = -1;
            window.Close();
        }
    }
}
