using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PopupServiceBack.Interfaces
{
    public interface IDialogService
    {
        string OpenTextDialog(string label, string name = null);
    }
}
