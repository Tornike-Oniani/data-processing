using PopupServiceBack.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PopupServiceBack.Interfaces
{
    public interface IWindowService
    {
        void OpenWindow(WindowViewModel viewModel);
    }
}
