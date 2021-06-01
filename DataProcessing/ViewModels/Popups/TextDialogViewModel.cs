using DataProcessing.Utils;
using DataProcessing.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class TextDialogViewModel : BaseViewModel
    {
        private IWindow window;

        // Public properties
        public string Label { get; set; }
        public string Input { get; set; }
        public string DialogResult { get; set; }

        // Commands
        public ICommand OkCommand { get; set; }
        public ICommand CancelCommand { get; set; }

        // Constructor
        public TextDialogViewModel(string label, IWindow window)
        {
            // Init
            this.Label = label;
            this.window = window;

            // Initialize commands
            OkCommand = new RelayCommand(Ok);
            CancelCommand = new RelayCommand(Cancel);
        }

        // Command actions
        public void Ok(object input = null)
        {
            DialogResult = Input;
            window.Close();
        }
        public void Cancel(object input = null)
        {
            DialogResult = null;
            window.Close();
        }
    }
}
