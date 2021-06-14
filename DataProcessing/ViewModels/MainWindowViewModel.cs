using DataProcessing.Classes;
using DataProcessing.Models;
using DataProcessing.Repositories;
using DataProcessing.Utils;
using DataProcessing.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace DataProcessing.ViewModels
{
    class MainWindowViewModel : BaseViewModel
    {
        // Private attributes
        private BaseViewModel _selectedViewModel;
        private bool _isHomeChecked;
        private bool _isWorkfileChecked;
        private bool _isWorking;
        private string _workLabel;

        // Public properties
        public BaseViewModel SelectedViewModel
        {
            get { return _selectedViewModel; }
            set { _selectedViewModel = value; OnPropertyChanged("SelectedViewModel"); }
        }
        public bool IsHomeChecked
        {
            get { return _isHomeChecked; }
            set { _isHomeChecked = value; OnPropertyChanged("IsHomeChecked"); }
        }
        public bool IsWorkfileChecked
        {
            get { return _isWorkfileChecked; }
            set { _isWorkfileChecked = value; OnPropertyChanged("IsWorkfileChecked"); }
        }
        public bool IsWorking
        {
            get { return _isWorking; }
            set { _isWorking = value; OnPropertyChanged("IsWorking"); }
        }
        public string WorkLabel
        {
            get { return _workLabel; }
            set { _workLabel = value; OnPropertyChanged("WorkLabel"); }
        }


        // Commands
        public ICommand UpdateViewCommand { get; set; }

        // Constructor
        public MainWindowViewModel()
        {
            // Init
            Services.GetInstance().SetWorkStatus = SetWorkStatus;
            Services.GetInstance().UpdateWorkStatus = UpdateWorkStatus;

            // Initialize commands
            UpdateViewCommand = new UpdateViewCommand(Navigate);
            ((UpdateViewCommand)UpdateViewCommand).OnChangeView += SetViewCheck;

            UpdateViewCommand.Execute(ViewType.Home);

            // Global error hanlder
            //if (Application.Current != null)
            //{
            //    Application.Current.DispatcherUnhandledException += (s, a) =>
            //    {
            //        // 2. Generic unhandled exceptions
            //        MessageBox.Show($"{a.Exception.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            //        a.Handled = true;
            //    };
            //}
        }

        // Command actions
        public void Navigate(BaseViewModel viewModel)
        {
            this.SelectedViewModel = viewModel;
        }

        // Event subscribers
        private void SetViewCheck(ViewType viewType)
        {
            switch (viewType)
            {
                case ViewType.Home: IsWorkfileChecked = false; IsHomeChecked = true; break;
                case ViewType.WorkfileEditor: IsHomeChecked = false; IsWorkfileChecked = true; break;
                default: break;
            }
        }

        // Private helpers
        private void SetWorkStatus(bool status)
        {
            this.IsWorking = status;
            this.WorkLabel = "Working...";
        }
        private void UpdateWorkStatus(string workLabel)
        {
            this.WorkLabel = workLabel;
        }
    }
}
