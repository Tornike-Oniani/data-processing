using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DataProcessing.Classes.Calculate
{
    internal class UserSelectedOptions
    {
        public string SelectedTimeMark { get; set; }
        public string SelectedRecordingType { get; set; }
        public int ClusterSparationTime { get; set; }
        public Dictionary<string, int[]> FrequencyRanges { get; set; }
        public List<SpecificCriteria> Criterias { get; set; }
    }
}
