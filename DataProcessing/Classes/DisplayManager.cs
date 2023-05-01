using DataProcessing.Models;
using DataProcessing.Utils;
using DataProcessing.ViewModels;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.Classes
{
    class DisplayManager : BaseViewModel
    {
        #region Property fields
        private string _selectedSheet;
        #endregion

        #region Properties
        public ObservableCollection<TimeStamp> Items { get; set; }
        public TimeStamp SelectedRow { get; set; }
        public List<TimeStamp> SelectedRows { get; set; }
        public List<string> Sheets { get; set; }

        public string SelectedSheet
        {
            get { return _selectedSheet; }
            set 
            { 
                _selectedSheet = value; 
                OnPropertyChanged("SelectedSheet");
                PopulateCommand.Execute(null);
            }
        }
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

            // Command initialization
            PopulateCommand = new RelayCommand(Populate);

            // We set the seet after initializing PopulateCommand because in setter we invoke it
            SelectedSheet = Sheets[0];
        }
        #endregion

        #region Public methods
        public Dictionary<string, List<TimeStamp>> GetDataForAllSheets()
        {
            Dictionary<string, List<TimeStamp>> result = new Dictionary<string, List<TimeStamp>>();

            List<TimeStamp> sheetData = new List<TimeStamp>();
            foreach (string sheet in Sheets)
            {
                sheetData = TimeStamp.Find(int.Parse(sheet.Substring(sheet.Length - 1)));
                result.Add(sheet, sheetData);
            }

            return result;
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
