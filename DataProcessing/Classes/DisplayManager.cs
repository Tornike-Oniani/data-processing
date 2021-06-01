using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.Classes
{
    class DisplayManager
    {
        // Properties
        public DataSample SelectedRow { get; set; }
        public ObservableCollection<DataSample> Items { get; set; }

        // Commands
        public ICommand PopulateCommand { get; set; }

        // Constructor
        public DisplayManager()
        {
            //Init
            Items = new ObservableCollection<DataSample>();

            // Command initialization
            PopulateCommand = new RelayCommand(Populate);
        }

        // Command actions
        public async void Populate(object input = null)
        {
            List<DataSample> items = new List<DataSample>();

            Services.GetInstance().SetWorkStatus(true);
            await Task.Run(() =>
            {
                items = DataSample.Find();
            });
            Services.GetInstance().SetWorkStatus(false);

            PopulateCollection(items);
        }

        // Private helpers
        private void PopulateCollection(List<DataSample> items)
        {
            Items.Clear();
            foreach (DataSample item in items)
            {
                Items.Add(item);
            }
        }
    }
}
