using DataProcessing.Classes;
using DataProcessing.Models;
using DataProcessing.Repositories;
using DataProcessing.Utils;
using DataProcessing.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class HomeViewModel : BaseViewModel
    {
        // Private attributes
        private ICommand updateViewCommand;
        private List<Workfile> _workfiles;
        private Workfile _selectedWorkfile;

        // Public properties
        public List<Workfile> Workfiles
        {
            get { return _workfiles; }
            set { _workfiles = value; OnPropertyChanged("Workfiles"); }
        }
        public Workfile SelectedWorkfile
        {
            get { return _selectedWorkfile; }
            set { _selectedWorkfile = value; OnPropertyChanged("SelectedWorkfile"); }
        }

        // Commands
        public ICommand ImportExcelCommand { get; set; }
        public ICommand OpenWorkfileCommand { get; set; }
        public ICommand DeleteWorkfileCommand { get; set; }
        public ICommand RenameWorkfileCommand { get; set; }

        // Constuctor
        public HomeViewModel(UpdateViewCommand updateViewCommand)
        {
            // Init
            this.updateViewCommand = updateViewCommand;
            Workfiles = WorkfileManager.GetInstance().GetWorkfiles();

            // Initialize commands
            ImportExcelCommand = new RelayCommand(ImportExcel);
            OpenWorkfileCommand = new RelayCommand(OpenWorkfile);
            DeleteWorkfileCommand = new RelayCommand(DeleteWorkfile);
            RenameWorkfileCommand = new RelayCommand(RenameWorkfile);
        }

        // Command actions
        public async void ImportExcel(object input = null)
        {
            // 1. Select file to import
            string file = Services.GetInstance().BrowserService.OpenFileDialog("", "Excel Files|*.xls;*.xlsx;*.xlsm");
            if (file == null) { return; }

            // 2. Create new workbook to import into
            string name = Services.GetInstance().DialogService.OpenTextDialog("Name:");
            if (name == null) { return; }

            WorkfileManager.GetInstance().CreateWorkfile(name);
            // TEMPORARY
            Workfile workfile = new Workfile() { Name = name };
            WorkfileManager.GetInstance().SelectedWorkFile = workfile;

            // 3. Import data
            await new ExcelManager().ImportFromExcel(file);

            // 4. Refresh Workfile list
            Workfiles = WorkfileManager.GetInstance().GetWorkfiles();
        }
        public void OpenWorkfile(object input = null)
        {
            if (SelectedWorkfile == null) { return; }
            updateViewCommand.Execute(ViewType.WorkfileEditor);
            WorkfileManager.GetInstance().SelectedWorkFile = SelectedWorkfile;
        }
        public void DeleteWorkfile(object input = null)
        {
            new WorkfileRepo().Delete(SelectedWorkfile);
            Workfiles = WorkfileManager.GetInstance().GetWorkfiles();
        }
        public void RenameWorkfile(object input = null)
        {
            string name = Services.GetInstance().DialogService.OpenTextDialog("Name:");
            if (String.IsNullOrEmpty(name)) { return; }
            string oldName = SelectedWorkfile.Name;
            SelectedWorkfile.Name = name;
            new WorkfileRepo().Update(SelectedWorkfile, oldName);
            Workfiles = WorkfileManager.GetInstance().GetWorkfiles();
        }
    }
}
