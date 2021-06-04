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
        public ObservableCollection<DataSample> Items { get; set; }
        public DataSample SelectedRow { get; set; }
        public List<DataSample> SelectedRows { get; set; }

        // Commands
        public ICommand PopulateCommand { get; set; }
        public ICommand TestCommand { get; set; }

        // Constructor
        public DisplayManager()
        {
            //Init
            Items = new ObservableCollection<DataSample>();
            SelectedRows = new List<DataSample>();

            // Command initialization
            PopulateCommand = new RelayCommand(Populate);
            TestCommand = new RelayCommand(Test);
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
        public void Test(object input = null)
        {
            foreach (var item in SelectedRows)
            {
                Console.WriteLine(item.AT);
            }
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
