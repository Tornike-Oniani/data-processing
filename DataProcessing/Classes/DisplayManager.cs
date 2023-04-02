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
        #region Properties
        public ObservableCollection<TimeStamp> Items { get; set; }
        public TimeStamp SelectedRow { get; set; }
        public List<TimeStamp> SelectedRows { get; set; }
        public List<string> Sheets { get; set; }
        public string SelectedSheet { get; set; }
        #endregion

        #region Commands
        public ICommand PopulateCommand { get; set; }
        #endregion

        #region Constructors
        public DisplayManager()
        {
            //Init
            Items = new ObservableCollection<TimeStamp>();
            SelectedRows = new List<TimeStamp>();
            Sheets = new List<string>();
            for (int i = 0; i < WorkfileManager.GetInstance().SelectedWorkFile.Sheets; i++)
            {
                Sheets.Add("Sheet" + (i + 1));
            }
            SelectedSheet = Sheets[0];

            // Command initialization
            PopulateCommand = new RelayCommand(Populate);
            PopulateCommand.Execute(null);
        }
        #endregion

        #region Command actions
        public async void Populate(object input = null)
        {
            List<TimeStamp> items = new List<TimeStamp>();

            Services.GetInstance().SetWorkStatus(true);
            await Task.Run(() =>
            {
                items = TimeStamp.Find(int.Parse(SelectedSheet.Substring(SelectedSheet.Length - 1)));
            });
            Services.GetInstance().SetWorkStatus(false);

            PopulateCollection(items);
        }
        #endregion

        #region Private helpers
        private void PopulateCollection(List<TimeStamp> items)
        {
            Items.Clear();
            foreach (TimeStamp item in items)
            {
                Items.Add(item);
            }
        }
        #endregion
    }
}
