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
    class WindowService : IWindowService
    {
        public void OpenWindow(BaseWindowViewModel viewModel)
        {
            Window window = new GenericWindow();
            viewModel.Window = (IWindow)window;
            window.DataContext = viewModel;
            window.Owner = Application.Current.MainWindow;
            window.ShowDialog();
        }
    }
}
