using DataProcessing.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Models
{
    class TimeStamp
    {
        private TimeStampRepo repo = new TimeStampRepo();
        private long TimeTicks;
        private long _timeDifference;

        // Properties
        public int Id { get; set; }
        public TimeSpan Time { get { return new TimeSpan(TimeTicks); } set { TimeTicks = value.Ticks; } }
        public TimeSpan TimeDifference { get { return new TimeSpan(_timeDifference); } set { _timeDifference = value.Ticks; } }
        public double TimeDifferenceInDouble { get; set; }
        public int TimeDifferenceInSeconds { get; set; }
        public int State { get; set; }
        public bool IsMarker { get; set; }
        public bool IsTimeMarked { get; set; }

        // Blank constructor
        public TimeStamp()
        {

        }

        // For cloning purposes (violates encapsulation though)
        public void SetTicks(long timeTicks)
        {
            this.TimeTicks = timeTicks;
        }
        public TimeStamp Clone()
        {
            TimeStamp cloned = new TimeStamp();
            cloned.SetTicks(this.TimeTicks);
            cloned.State = this.State;
            return cloned;
        }

        // Database functions
        public void Save()
        {
            repo.Create(this);
        }
        public static void SaveMany(List<TimeStamp> records, int sheetNumber)
        {
            new TimeStampRepo().CreateMany(records, sheetNumber);
        }
        public void Update()
        {
            repo.Update(this);
        }
        public static List<TimeStamp> Find(int sheetNumber)
        {
            return new TimeStampRepo().Find(sheetNumber);
        }

        // Private helpers
        public void CalculateStatsWhenMany(TimeStamp previous)
        {
            CalculateB(previous);
            CalculateC();
            CalculateD();
        }
        private void CalculateB(TimeStamp previous)
        {
            if (Time < previous.Time)
            {
                TimeDifference = Time + new TimeSpan(24, 0, 0) - previous.Time;
                return;
            }

            TimeDifference = Time - previous.Time;
        }
        private void CalculateC()
        {
            TimeDifferenceInDouble = TimeDifference.TotalDays;
        }
        private void CalculateD()
        {
            TimeDifferenceInSeconds = (int)Math.Round(TimeDifferenceInDouble * 86400);
        }
    }
}
