using DataProcessing.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    class Services
    {
        // Singleton Implementation
        private Services() { }
        private static Services _instance;
        public static Services GetInstance()
        {
            if (_instance == null)
            {
                _instance = new Services();
            }
            return _instance;
        }

        public IDialogService DialogService { get; set; }
        public IBrowserService BrowserService { get; set; }
        public IWindowService WindowService { get; set; }
        public Action<bool> SetWorkStatus { get; set; }
        public Action<string> UpdateWorkStatus { get; set; }
    }
}
