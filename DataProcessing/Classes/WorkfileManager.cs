using DataProcessing.Models;
using DataProcessing.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    class WorkfileManager : ObservableObject
    {
        // Singleton Implementation
        private WorkfileManager() { }
        private static WorkfileManager _instance;
        public static WorkfileManager GetInstance()
        {
            if (_instance == null)
            {
                _instance = new WorkfileManager();
            }
            return _instance;
        }

        // Private attributes
        private Workfile _selectedWorkFile;

        // Properties
        public Workfile SelectedWorkFile
        {
            get { return _selectedWorkFile; }
            set 
            { 
                _selectedWorkFile = value; 
                OnPropertyChanged("SelectedWorkFile");
                OnWorkfileChanged?.Invoke(SelectedWorkFile);
            }
        }

        // Events
        public event Action<Workfile> OnWorkfileChanged;

        // Database operations
        public void CreateWorkfile(string name) { new WorkfileRepo().Create(name); }
        public List<Workfile> GetWorkfiles() { return new WorkfileRepo().Find(); }
        public void UpdateWorkgile(Workfile workfile, string oldName) { new WorkfileRepo().Update(workfile, oldName); }
        public void DeleteWorkgile(Workfile workfile) { new WorkfileRepo().Delete(workfile); }

    }
}
