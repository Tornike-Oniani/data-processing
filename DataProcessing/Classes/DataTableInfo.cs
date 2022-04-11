using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    /// <summary>
    /// Extension of DataTable that also depicts addition information for displaying it to excel file
    /// NOTE: might be better to use inheritance here
    /// </summary>
    internal class DataTableInfo
    {
        // Table containing all the data
        public DataTable Table { get; set; }
        // Does the table data contain information for the whole sample region
        public bool IsTotal { get; set; }
        // Do we need to visualise table as chart
        public bool DisplayChart { get; set; }
    }
}
