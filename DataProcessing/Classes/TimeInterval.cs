using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

namespace DataProcessing.Classes
{
    internal class TimeInterval
    {
        #region Public properties
        public TimeSpan From { get; set; }
        public TimeSpan Till { get; set; }
        #endregion

        #region Public methods
        public bool IsCorrect()
        {
            return From < Till;
        }
        public static bool IsBetweenTimeInterval(TimeInterval interval, TimeSpan from, TimeSpan to)
        {
            return interval.From >= from && interval.Till <= to;
        }
        #endregion
    }
}
