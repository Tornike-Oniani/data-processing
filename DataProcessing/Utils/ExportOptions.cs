using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Utils
{
    class ExportOptions
    {
        public float TimeMark { get; set; }
        public int MaxStates { get; set; }
        public TimeSpan From { get; set; }
        public TimeSpan Till { get; set; }
        public List<SpecificCriteria> Criterias { get; set; }

        public List<SpecificCriteria> GetExistentCriterias()
        {
            return Criterias.Where(c => c.Value != null).ToList();
        }
    }
}
