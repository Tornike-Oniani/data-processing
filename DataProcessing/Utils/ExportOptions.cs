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
        public Dictionary<int, int> StateAndCriteria { get; set; } = new Dictionary<int, int>();
    }
}
