using DataProcessing.Classes;
using DataProcessing.Models;
using DataProcessing.Repositories;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class HomeViewModel : BaseViewModel
    {
        // Private attributes
        private ICommand updateViewCommand;
        private Workfile _selectedWorkfile;
        private string _search;

        // Public properties
        public ObservableCollection<Workfile> Workfiles { get; set; }
        public CollectionViewSource _workfilesCollection { get; set; }
        public ICollectionView WorkfilesCollection { get { return _workfilesCollection.View; } }
        public Workfile SelectedWorkfile
        {
            get { return _selectedWorkfile; }
            set { _selectedWorkfile = value; OnPropertyChanged("SelectedWorkfile"); }
        }
        public string Search
        {
            get { return _search; }
            set 
            { 
                _search = value;
                _workfilesCollection.View.Refresh();
                OnPropertyChanged("Search");
            }
        }


        // Commands
        public ICommand ImportExcelCommand { get; set; }
        public ICommand OpenWorkfileCommand { get; set; }
        public ICommand DeleteWorkfileCommand { get; set; }
        public ICommand RenameWorkfileCommand { get; set; }
        public ICommand ClearSearchCommand { get; set; }

        // Constuctor
        public HomeViewModel(UpdateViewCommand updateViewCommand)
        {
            // Init
            this.updateViewCommand = updateViewCommand;
            Workfiles = new ObservableCollection<Workfile>();
            _workfilesCollection = new CollectionViewSource();
            _workfilesCollection.Source = Workfiles;
            _workfilesCollection.Filter += OnSearch;
            PopulateWorkfiles(WorkfileManager.GetInstance().GetWorkfiles());

            // Initialize commands
            ImportExcelCommand = new RelayCommand(ImportExcel);
            OpenWorkfileCommand = new RelayCommand(OpenWorkfile);
            DeleteWorkfileCommand = new RelayCommand(DeleteWorkfile);
            RenameWorkfileCommand = new RelayCommand(RenameWorkfile);
            ClearSearchCommand = new RelayCommand(ClearSearch);
        }

        private void OnSearch(object sender, FilterEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(Search))
            {
                e.Accepted = true;
                return;
            }

            e.Accepted = false;

            Workfile current = e.Item as Workfile;
            if (current.Name.ToUpper().Contains(Search.ToUpper()))
            {
                e.Accepted = true;
            }
        }

        // Command actions
        public async void ImportExcel(object input = null)
        {
            WorkfileManager workfileManager = WorkfileManager.GetInstance();

            // 1. Select file to import
            string file = Services.GetInstance().BrowserService.OpenFileDialog("", "Excel Files|*.xls;*.xlsx;*.xlsm");
            if (file == null) { return; }

            // 2. Create new workfile to import into
            string name = Services.GetInstance().DialogService.OpenTextDialog("Name:", Path.GetFileNameWithoutExtension(file));
            if (name == null) { return; }

            WorkfileManager.GetInstance().CreateWorkfile(new Workfile() { Name = name, ImportDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") });
            // TEMPORARY (MAYBE FIXED?)
            workfileManager.SelectedWorkFile = workfileManager.GetWorkfileByName(name);

            // 3. Check file for errors
            ExcelManager excelManager = new ExcelManager(new ExportOptions(), null, null, null, null, null, null, null, null, null);
            List<int> errorRows = new List<int>();
            try
            {
                
                errorRows = await excelManager.CheckExcelFile(file);
            }
            catch (Exception e)
            {
                Workfile wf = WorkfileManager.GetInstance().GetWorkfileByName(name);
                WorkfileManager.GetInstance().DeleteWorkfile(wf);
                throw e;
            }

            if (errorRows.Count > 0)
            {
                MessageBoxResult result = MessageBox.Show("There might be erorrs in the excel file, do you want to stop importing and highlight possible errors?\nYes - Stop import and highlight errors\nNo - import file", "Excel file check", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.Yes)
                {
                    await excelManager.HighlightExcelFileErrors(file, errorRows);
                    workfileManager.DeleteWorkfile(workfileManager.SelectedWorkFile);
                    return;
                }
            }

            // 4. Import data
            await excelManager.ImportFromExcel(file);

            // 5. Refresh Workfile list
            PopulateWorkfiles(WorkfileManager.GetInstance().GetWorkfiles());
        }
        public void OpenWorkfile(object input = null)
        {
            if (SelectedWorkfile == null) { return; }
            updateViewCommand.Execute(ViewType.WorkfileEditor);
            WorkfileManager.GetInstance().SelectedWorkFile = SelectedWorkfile;
        }
        public void DeleteWorkfile(object input = null)
        {
            MessageBoxResult dialogResult = MessageBox.Show($"Are you sure you want to delete \"{SelectedWorkfile.Name}\"?", "Delete workfile", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (dialogResult == MessageBoxResult.No) { return; }

            new WorkfileRepo().Delete(SelectedWorkfile);
            PopulateWorkfiles(WorkfileManager.GetInstance().GetWorkfiles());
        }
        public void RenameWorkfile(object input = null)
        {
            string oldName = SelectedWorkfile.Name;
            string name = Services.GetInstance().DialogService.OpenTextDialog("Name:", oldName);
            if (String.IsNullOrEmpty(name) || oldName == name) { return; }
            SelectedWorkfile.Name = name;
            new WorkfileRepo().Update(SelectedWorkfile, oldName);
            PopulateWorkfiles(WorkfileManager.GetInstance().GetWorkfiles());
        }
        public void ClearSearch(object input = null)
        {
            Search = null;
        }

        // Private helpers
        private void PopulateWorkfiles(List<Workfile> workfiles)
        {
            Workfiles.Clear();
            foreach (Workfile workfile in workfiles)
            {
                Workfiles.Add(workfile);
            }
        }
    } 
}
