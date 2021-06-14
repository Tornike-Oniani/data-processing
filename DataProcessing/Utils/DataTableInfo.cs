using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Utils
{
    class DataTableInfo
    {
        public DataTable Table { get; set; }
        public Tuple<int, int> HeaderIndexes { get; set; }
        public Tuple<int, int> PhasesIndexes { get; set; }
        public Tuple<int, int> CriteriaPhases { get; set; }
        public bool IsTotal { get; set; }
    }
}
