using DataProcessing.Classes;
using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class OpenWorkfileViewModel : BaseWindowViewModel
    {
        // Private attributes
        private Workfile _selectedWorkfile;

        // Public properties
        public List<Workfile> Workfiles { get; set; }
        public Workfile SelectedWorkfile
        {
            get { return _selectedWorkfile; }
            set { _selectedWorkfile = value; OnPropertyChanged("SelectedWorkFile"); }
        }

        // Command
        public ICommand OpenWorkfileCommand { get; set; }
        public ICommand CloseCommand { get; set; }

        // Constructor
        public OpenWorkfileViewModel()
        {
            // Init
            this.Title = "Open...";
            Workfiles = WorkfileManager.GetInstance().GetWorkfiles();

            // Initialize commands
            OpenWorkfileCommand = new RelayCommand(OpenWorkfile);
            CloseCommand = new RelayCommand(Close);
        }

        // Command actions
        public void OpenWorkfile(object input = null)
        {
            WorkfileManager.GetInstance().SelectedWorkFile = SelectedWorkfile;
            this.Window.Close();
        }
        public void Close(object input = null)
        {
            this.Window.Close();
        }
    }
}
