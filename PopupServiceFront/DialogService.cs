using PopupServiceBack;
using PopupServiceBack.Interfaces;
using PopupServiceBack.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PopupServiceFront
{
    public class DialogService : IDialogService
    {
        public string OpenTextDialog(string label, string name = null)
        {
            Window window = new GenericWindow();
            window.Owner = Application.Current.MainWindow;
            window.DataContext = new TextDialogViewModel(label, name, (IWindow)window);
            window.ShowDialog();
            return (window.DataContext as TextDialogViewModel).DialogResult;
        }
    }
}
