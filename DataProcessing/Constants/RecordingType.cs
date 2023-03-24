using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Constants
{
    internal static class RecordingType
    {
        public static readonly string ThreeStates = "1 - Paradoxical; 2 - Sleep; 3 - Wakefulness";
        public static readonly string TwoStates = "1 - Sleep; 2 - Wakefulness";
        public static readonly string TwoStatesWithBehavior = "1 - Sleep; 2 - Wakefulness + Behaviors (3 - Active; 4 - Passive;\n5 - Grooming; 6 - Eating; 7 - Water)";

        public static readonly Dictionary<string, int> MaxStates = new Dictionary<string, int>()
        {
            {ThreeStates, 3},
            {TwoStates, 2},
            {TwoStatesWithBehavior, 7}
        };

        public static Dictionary<int, string> GetThreeStatesDictionary()
        {
            return new Dictionary<int, string>()
            {
                { 3, "Wakefulness" },
                { 2, "Sleep" },
                { 1, "Paradoxical sleep" }
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
        public static Dictionary<int, string> GetTwoStatesWithBehaviorDictionary()
        {
            return new Dictionary<int, string>()
            {
                {1, "Sleep" },
                {2, "Wakefulness" },
                {3, "Active" },
                {4, "Passive" },
                {5, "Grooming" },
                {6, "Eating" },
                {7, "Water" },
            };
        }
    }
}
