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
        public void CreateWorkfile(Workfile workfile) { new WorkfileRepo().Create(workfile); }
        public List<Workfile> GetWorkfiles() { return new WorkfileRepo().Find(); }
        public void UpdateWorkfile(Workfile workfile, string oldName) { new WorkfileRepo().Update(workfile, oldName); }
        public void DeleteWorkfile(Workfile workfile) { new WorkfileRepo().Delete(workfile); this.SelectedWorkFile = null; }
        public Workfile GetWorkfileByName(string name) { return new WorkfileRepo().FindByName(name); }

    }
}
