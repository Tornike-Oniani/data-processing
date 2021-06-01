using DataProcessing.Classes;
using DataProcessing.Utils.Services;
using DataProcessing.ViewModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DataProcessing.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = new MainWindowViewModel();

            this.Loaded += (s, e) =>
            {
                Services.GetInstance().DialogService = new DialogService();
                Services.GetInstance().BrowserService = new BrowserService();
                Services.GetInstance().WindowService = new WindowService();
            };
        }
    }
}
