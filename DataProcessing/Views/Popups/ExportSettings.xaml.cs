using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    /// Interaction logic for ExportSettings.xaml
    /// </summary>
    public partial class ExportSettings : UserControl
    {
        public ExportSettings()
        {
            InitializeComponent();
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox txbThis = sender as TextBox;
            if (String.IsNullOrWhiteSpace(txbThis.Text)) { txbThis.Text = null; }
        }
    }
}
