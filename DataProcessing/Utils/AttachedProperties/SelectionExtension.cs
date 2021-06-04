using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace DataProcessing.Utils
{
    static class SelectionExtension
    {
        public static IList GetSelectedItems(DependencyObject d)
        {
            return (IList)d.GetValue(SelectedItemsProperty);
        }

        public static void SetSelectedItems(DependencyObject d, IList value)
        {
            d.SetValue(SelectedItemsProperty, value);
        }

        public static readonly DependencyProperty SelectedItemsProperty =
            DependencyProperty.RegisterAttached("SelectedItems", typeof(IList),
            typeof(SelectionExtension),
            new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault,
            new PropertyChangedCallback(OnSelectedItemsChanged)));

        private static void OnSelectedItemsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            DataGrid grid = (DataGrid)d;
            grid.SelectionChanged += DataGrid_SelectionChanged;
        }

        private static void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DataGrid grid = (DataGrid)sender;
            //Get list box's selected items.
            IEnumerable selectedItems = grid.SelectedItems;
            //Get list from model
            IList ModelSelectedItems = GetSelectedItems(grid);

            //Update the model
            ModelSelectedItems.Clear();

            if (grid.SelectedItems != null)
            {
                foreach (var item in grid.SelectedItems)
                    ModelSelectedItems.Add(item);
            }
            SetSelectedItems(grid, ModelSelectedItems);
        }
    }
}
