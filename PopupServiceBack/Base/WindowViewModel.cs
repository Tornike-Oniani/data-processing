using PopupServiceBack.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PopupServiceBack.Base
{
    public class WindowViewModel
    {
        public IWindow Window { get; set; }
        public string Title { get; set; }
    }
}
