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
        public static readonly string TwoStatesWithBehavior = "1 - Sleep + Behaviors (2 - Active; 3 - Passive; 4 - Grooming; 5 - Eating; 6 - Water";

        public static readonly Dictionary<string, int> MaxStates = new Dictionary<string, int>()
        {
            {ThreeStates, 3},
            {TwoStates, 2},
            {TwoStatesWithBehavior, 6}
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
        public static Dictionary<int, string> GetTwoStatesWithBehaviorDictionary()
        {
            return new Dictionary<int, string>()
            {
                {1, "Sleep" },
                {2, "Active" },
                {3, "Passive" },
                {4, "Grooming" },
                {5, "Eating" },
                {6, "Water" },
            };
        }
    }
}
