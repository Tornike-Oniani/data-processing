using DataProcessing.Utils.Interfaces;
using DataProcessing.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace DataProcessing.Utils.Services
{
    class DialogService : IDialogService
    {
        public string OpenTextDialog(string label)
        {
            Window window = new GenericWindow();
            window.Owner = Application.Current.MainWindow;
            window.DataContext = new TextDialogViewModel(label, (IWindow)window);
            window.ShowDialog();
            return (window.DataContext as TextDialogViewModel).DialogResult;
        }

        public int OpenRadioDialog()
        {
            Window window = new GenericWindow();
            window.Owner = Application.Current.MainWindow;
            window.DataContext = new RadioDialogViewModel((IWindow)window);
            window.ShowDialog();
            return (window.DataContext as RadioDialogViewModel).DialogResult;
        }
    }
}
