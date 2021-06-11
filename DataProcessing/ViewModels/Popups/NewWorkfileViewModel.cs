using DataProcessing.Classes;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class NewWorkfileViewModel : BaseWindowViewModel
    {
        // Private attributes
        private string _name;

        // Public properties
        public string Name
        {
            get { return _name; }
            set { _name = value; OnPropertyChanged("Name"); }
        }

        // Commands
        public ICommand CreateWorkfileCommand { get; set; }
        public ICommand CloseCommand { get; set; }

        // Constructor
        public NewWorkfileViewModel()
        {
            // Init
            this.Title = "Create...";

            // Initialize commands
            CreateWorkfileCommand = new RelayCommand(CreateWorkfile);
            CloseCommand = new RelayCommand(Close);
        }

        // Command actions
        public void CreateWorkfile(object input = null)
        {
            if (String.IsNullOrEmpty(Name)) throw new Exception("Workfile must have a name!");

            WorkfileManager.GetInstance().CreateWorkfile(new Models.Workfile() { Name = Name, ImportDate = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") });
            this.Window.Close();
        }
        public void Close(object input = null)
        {
            this.Window.Close();
        }
    }
}
