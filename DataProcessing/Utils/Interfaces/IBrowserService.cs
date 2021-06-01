using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Utils.Interfaces
{
    interface IBrowserService
    {
        string OpenFileDialog(string defaultEx, string filter);
    }
}
