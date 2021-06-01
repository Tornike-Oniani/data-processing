using System;
using System.Collections.Generic;
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
    /// Interaction logic for WorkfileView.xaml
    /// </summary>
    public partial class WorkfileEditor : UserControl
    {
        public WorkfileEditor()
        {
            InitializeComponent();
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (int.TryParse(e.Text, out _))
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void TextBox_TimeSpanInput(object sender, TextCompositionEventArgs e)
        {
            Regex timeSpanFormat = new Regex("^[0-9:]+$", RegexOptions.None);
            if (timeSpanFormat.IsMatch(e.Text))
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void DataGrid_LostFocus(object sender, RoutedEventArgs e)
        {
            DataGrid dg = sender as DataGrid;
            dg.UnselectAll();
        }

        private void DataGrid_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is ScrollViewer)
            {
                ((DataGrid)sender).UnselectAll();
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            //this.myDataGrid.SelectionMode = DataGridSelectionMode.Extended;
            //this.myDataGrid.SelectAllCells();
            //this.myDataGrid.ClipboardCopyMode = DataGridClipboardCopyMode.ExcludeHeader;
            //ApplicationCommands.Copy.Execute(null, this.myDataGrid);
            //this.myDataGrid.UnselectAllCells();
            //this.myDataGrid.SelectionMode = DataGridSelectionMode.Single;
            //String result = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);
            //StreamWriter sw = new StreamWriter("wpfdata.csv");
            //sw.WriteLine(result);
            //sw.Close();
            //Process.Start("wpfdata.csv");
        }
    }
}
