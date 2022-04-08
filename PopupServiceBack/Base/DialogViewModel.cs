using PopupServiceBack.Interfaces;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PopupServiceBack.Base
{
    public class DialogViewModel : INotifyPropertyChanged
    {
        protected IWindow window;

        public string Title { get; set; }
        public string Text { get; set; }
        public bool UserDialogResult { get; set; }

        public DialogViewModel(string text, string title, IWindow window)
        {
            this.Title = title;
            this.Text = text;
            this.window = window;
        }

        public void SetDialogResult(bool result)
        {
            this.UserDialogResult = result;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged(string propertyName)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
