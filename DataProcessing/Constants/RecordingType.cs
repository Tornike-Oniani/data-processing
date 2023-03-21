using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Constants
{
    internal static class RecordingType
    {
        public static readonly string ThreeStates = "3 - Wakefulness; 2 - Sleep; 1 - Paradoxical";
        public static readonly string TwoStates = "2 - Wakefulness; 1 - Sleep";

        public static readonly Dictionary<string, int> MaxStates = new Dictionary<string, int>()
        {
            {ThreeStates, 3},
            {TwoStates, 2}
        };

        public static Dictionary<int, string> GetThreeStatesDictionary()
        {
            return new Dictionary<int, string>()
            {
                { 1, "Paradoxical Sleep" },
                { 2, "Sleep" },
                { 3, "Wakefulness" }
            };
        }
        public static Dictionary<int, string> GetTwoStatesDictionary()
        {
            return new Dictionary<int, string>()
            {
                {1, "Sleep" },
                {2, "Wakefulness" }
            };
        }
    }
}
