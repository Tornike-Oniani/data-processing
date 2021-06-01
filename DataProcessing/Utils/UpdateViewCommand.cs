using DataProcessing.Utils.Interfaces;
using DataProcessing.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.Utils
{
    class UpdateViewCommand : ICommand
    {
        private Action<BaseViewModel> navigate;

        public UpdateViewCommand(Action<BaseViewModel> navigate)
        {
            this.navigate = navigate;
        }

        public event EventHandler CanExecuteChanged;
        public event Action<ViewType> OnChangeView;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            ViewType viewType = (ViewType)parameter;
            switch (viewType)
            {
                case ViewType.Home:
                    navigate(new HomeViewModel(this));
                    break;
                case ViewType.WorkfileEditor:
                    navigate(new WorkfileEditorViewModel());
                    break;
            }

            OnChangeView.Invoke(viewType);
        }
    }
}
