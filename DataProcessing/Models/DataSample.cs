using DataProcessing.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Models
{
    class DataSample
    {
        private SampleRepo repo = new SampleRepo();
        private long A;
        private long B;
        private int Counter;

        // Properties
        public int Id { get; set; }
        public TimeSpan AT { get { return new TimeSpan(A); } set { A = value.Ticks; } }
        public TimeSpan BT { get { return new TimeSpan(B); } set { B = value.Ticks; } }
        public double C { get; set; }
        public int D { get; set; }
        public int State { get; set; }
        public bool IsMarker { get; set; }
        public bool IsTimeMarked { get; set; }

        // Blank constructor
        public DataSample()
        {

        }

        // Database functions
        public void Save()
        {
            CalculateStats(true);
            repo.Create(this);
        }
        public static void SaveMany(List<DataSample> samples)
        {
            // Calculate
            //for (int i = 1; i < samples.Count; i++)
            //{
            //    samples[i].CalculateStatsWhenMany(samples[i - 1]);
            //}

            new SampleRepo().CreateMany(samples);
        }
        public void Update()
        {
            CalculateStats(false);
            repo.Update(this);
            DataSample linkedRecord = repo.GetLinkedRecord(this);

            if (linkedRecord == null) { return; }

            linkedRecord.CalculateStats(false);
            repo.Update(linkedRecord);

        }
        public static List<DataSample> Find()
        {
            return new SampleRepo().Find();
        }

        // Private helpers
        public void CalculateStats(bool isNew)
        {
            DataSample previous = isNew ? repo.GetLastRecord() : repo.GetPreviousRecord(Counter);
            if (previous == null) return;

            CalculateB(previous);
            CalculateC();
            CalculateD();
        }
        public void CalculateStatsWhenMany(DataSample previous)
        {
            CalculateB(previous);
            CalculateC();
            CalculateD();
        }
        private void CalculateB(DataSample previous)
        {
            if (AT < previous.AT)
            {
                BT = AT + new TimeSpan(24, 0, 0) - previous.AT;
                return;
            }

            BT = AT - previous.AT;
        }
        private void CalculateC()
        {
            C = BT.TotalDays;
        }
        private void CalculateD()
        {
            D = (int)Math.Round(C * 86400);
        }
    }
}
