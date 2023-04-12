using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    internal class ExcelSheetErrors
    {
        #region Public properties
        public List<int> MainDataErrorRows { get; set; }
        public Dictionary<int, List<int>> BehaviorErrorRows { get; set; }
        #endregion

        #region Public methods
        public void AddMainDataErrorRow(int rowIndex)
        {
            MainDataErrorRows.Add(rowIndex);
        }
        public void AddBehaviorErrorRow(int behavior, int rowIndex)
        {
            // 1. Initialize error list of behaviour if its the first one
            if (!BehaviorErrorRows.ContainsKey(behavior))
            {
                BehaviorErrorRows.Add(behavior, new List<int>());
            }

            BehaviorErrorRows[behavior].Add(rowIndex);
        }
        #endregion
    }
}
