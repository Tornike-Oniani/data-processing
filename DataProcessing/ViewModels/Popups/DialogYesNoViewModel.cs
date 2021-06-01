﻿using DataProcessing.Utils;
using DataProcessing.Utils.Interfaces;
using DataProcessing.ViewModels.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.ViewModels.Popups
{
    class DialogYesNoViewModel : BaseDialogViewModel
    {
        public ICommand YesCommand { get; set; }
        public ICommand NoCommand { get; set; }

        public DialogYesNoViewModel(string text, string title, IWindow window) : base(text, title, window)
        {
            YesCommand = new RelayCommand(Yes);
            NoCommand = new RelayCommand(No);
        }

        public void Yes(object input)
        {
            this.SetDialogResult(true);
            window.Close();
        }

        public void No(object input)
        {
            this.SetDialogResult(false);
            window.Close();
        }
    }
}
