using PopupServiceBack.Base;
using PopupServiceBack.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace PopupServiceBack.ViewModels
{
    public class TextDialogViewModel : WindowViewModel
    {
        // Public properties
        public string Label { get; set; }
        public string Input { get; set; }
        public string DialogResult { get; set; }

        // Commands
        public ICommand OkCommand { get; set; }
        public ICommand CancelCommand { get; set; }

        // Constructor
        public TextDialogViewModel(string label, string input, IWindow window)
        {
            // Init
            this.Title = label;
            this.Label = label;
            this.Window = window;
            this.Input = input;

            // Initialize commands
            OkCommand = new RelayCommand(Ok);
            CancelCommand = new RelayCommand(Cancel);
        }

        // Command actions
        public void Ok(object input = null)
        {
            DialogResult = Input;
            Window.Close();
        }
        public void Cancel(object input = null)
        {
            DialogResult = null;
            Window.Close();
        }
    }
}
