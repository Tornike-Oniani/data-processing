using PopupServiceBack.Base;
using PopupServiceBack.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace PopupServiceFront
{
    public class WindowService : IWindowService
    {
        public void OpenWindow(WindowViewModel viewModel)
        {
            Window window = new GenericWindow();
            viewModel.Window = (IWindow)window;
            window.DataContext = viewModel;
            window.Owner = Application.Current.MainWindow;
            window.ShowDialog();
        }
    }
}
