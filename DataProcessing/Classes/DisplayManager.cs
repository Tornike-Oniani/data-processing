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
        public ObservableCollection<TimeStamp> Items { get; set; }
        public TimeStamp SelectedRow { get; set; }
        public List<TimeStamp> SelectedRows { get; set; }

        // Commands
        public ICommand PopulateCommand { get; set; }
        public ICommand TestCommand { get; set; }

        // Constructor
        public DisplayManager()
        {
            //Init
            Items = new ObservableCollection<TimeStamp>();
            SelectedRows = new List<TimeStamp>();

            // Command initialization
            PopulateCommand = new RelayCommand(Populate);
            TestCommand = new RelayCommand(Test);
        }

        // Command actions
        public async void Populate(object input = null)
        {
            List<TimeStamp> items = new List<TimeStamp>();

            Services.GetInstance().SetWorkStatus(true);
            await Task.Run(() =>
            {
                items = TimeStamp.Find();
            });
            Services.GetInstance().SetWorkStatus(false);

            PopulateCollection(items);
        }
        public void Test(object input = null)
        {
            foreach (var item in SelectedRows)
            {
                Console.WriteLine(item.Time);
            }
        }

        // Private helpers
        private void PopulateCollection(List<TimeStamp> items)
        {
            Items.Clear();
            foreach (TimeStamp item in items)
            {
                Items.Add(item);
            }
        }
    }
}
