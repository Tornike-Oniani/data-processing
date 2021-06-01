using DataProcessing.Utils.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.ViewModels
{
    class BaseWindowViewModel : BaseViewModel
    {
        public IWindow Window { get; set; }
        public string Title { get; set; }
    }
}
