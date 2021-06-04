using DataProcessing.Models;
using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace DataProcessing.Classes
{
    class EntryManager : ObservableObject
    {
        private string _timeStamp;
        private bool _isEntryFocused;
        private ICommand populate;

        // Properties
        public string TimeStamp
        {
            get { return _timeStamp; }
            set { _timeStamp = value; OnPropertyChanged("TimeStamp"); }
        }
        public bool IsEntryFocused
        {
            get { return _isEntryFocused; }
            set { _isEntryFocused = value; OnPropertyChanged("IsEntryFocused"); }
        }

        // Commands
        public ICommand AddCommand { get; set; }

        // Constructor
        public EntryManager(ICommand populate)
        {
            // Command initialization
            AddCommand = new RelayCommand(Add);
            this.populate = populate;
        }

        // Command actions
        public async void Add(object input = null)
        {
            if (String.IsNullOrWhiteSpace(TimeStamp)) { IsEntryFocused = true; throw new Exception("TimeStamp can not be empty!"); }
            if (TimeStamp.Length != 7) { IsEntryFocused = true; throw new Exception("TimeStamp has to be 7 characters long!"); }

            Tuple<TimeSpan, int> timeAndState = GetTimeAndState(TimeStamp);

            if (timeAndState.Item1.Days != 0) { IsEntryFocused = true; throw new Exception("TimeStamp can not have more than 24 hours!"); }

            Services.GetInstance().SetWorkStatus(true);

            await Task.Run(() =>
            {
                TimeStamp sample = new TimeStamp() { Time = timeAndState.Item1, State = timeAndState.Item2 };
                sample.Save();
            });

            Services.GetInstance().SetWorkStatus(false);

            TimeStamp = null;
            IsEntryFocused = true;

            populate.Execute(null);
        }

        // Private helpers
        private Tuple<TimeSpan, int> GetTimeAndState(string timeStamp)
        {
            int step = 0;
            int hour = int.Parse(Cut(timeStamp, ref step));
            int minutes = int.Parse(Cut(timeStamp, ref step));
            int seconds = int.Parse(Cut(timeStamp, ref step));
            int state = int.Parse(timeStamp.Substring(timeStamp.Length - 1, 1));
            TimeSpan time = new TimeSpan(hour, minutes, seconds);
            return new Tuple<TimeSpan, int>(time, state);
        }
        private string Cut(string input, ref int step)
        {
            string res = input.Substring(step, 2);
            step += 2;
            return res;
        }
    }
}
