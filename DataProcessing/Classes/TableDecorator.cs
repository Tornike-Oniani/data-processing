using DataProcessing.Classes.Export;
using DataProcessing.Models;
using System;
using System.Collections.Generic;

namespace DataProcessing.Classes
{
    internal class TableDecorator
    {
        private int _maxStates;

        public TableDecorator(int maxStates)
        {
            this._maxStates = maxStates;
        }

        public ExcelTable DecorateRawData(object[,] data, List<TimeStamp> timeStamps, int timeMarkInSeconds)
        {
            ExcelTable table = new ExcelTable(data);
            // Default hour distinction colors
            int time = 0;
            TimeStamp cur;
            for (int i = 0; i < timeStamps.Count; i++)
            {
                cur = timeStamps[i];
                time += cur.TimeDifferenceInSeconds;
                // If timestamp was added programatically for episode divison color it dark green
                if (cur.IsTimeMarked) { table.AddColor("DarkGreen", new ExcelRange(i, 0, i, 4)); }
                // If timestamp was added programatically for 10am purposes color it yellow
                if (cur.IsMarker) { table.AddColor("Yellow", new ExcelRange(i, 0, i, 4)); }
                // If we naturally reached the end of episode color it green
                if (time == timeMarkInSeconds)
                {
                    if (!cur.IsTimeMarked && !cur.IsMarker)
                    {
                        table.AddColor("Green", new ExcelRange(i, 0, i, 4));
                    }
                    time = 0;
                }
                // If we passed natural end of episode without marking it throw exception
                if (time > timeMarkInSeconds) { throw new Exception("Incorrect time mark division."); }
            }

            return table;
        }
        public ExcelTable DecorateLatencyTable(object[,] data)
        {
            ExcelTable table = new ExcelTable(data);

            // Header
            table.AddColor("Orange", new ExcelRange(0, 0, 0, 0));
            table.AddColor("Blue", new ExcelRange(1, 0, 1, _maxStates - 2));

            table.SetHeaderRange(1, 0, 1, _maxStates - 2);

            return table;
        }
        public ExcelTable DecorateStatTable(object[,] data, int criteriaNumber)
        {
            StatTable table = new StatTable(data);

            // Header
            table.AddColor("Orange", new ExcelRange(0, 0, 0, 4));
            // Phases
            table.AddColor("Blue", new ExcelRange(1, 0, _maxStates, 0));
            // Specific criterias
            if (criteriaNumber != 0)
            {
                table.AddColor("Red", new ExcelRange(_maxStates + 1, 0, _maxStates + criteriaNumber, 0));
            }

            table.SetHeaderRange(0, 1, 0, 4);

            return table;
        }
        public ExcelTable DecorateStatTableTotal(object[,] data, int criteriaNumber)
        {
            StatTable table = new StatTable(data);

            // Header
            table.AddColor("DarkOrange", new ExcelRange(0, 0, 0, 4));
            // Phases
            table.AddColor("DarkBlue", new ExcelRange(1, 0, _maxStates + 1, 0));
            // Specific criterias
            if (criteriaNumber != 0)
            {
                table.AddColor("DarkRed", new ExcelRange(_maxStates + 2, 0, _maxStates + 1 + criteriaNumber, 0));
            }

            table.SetHeaderRange(0, 1, 0, 4);

            return table;
        }
        public ExcelTable DecorateGraphTable(object[,] data, bool hasChart)
        {
            ExcelTable table = hasChart ? new GraphTableWithChart(data) : new ExcelTable(data);
            int columnCount = data.GetLength(1);
            // Header
            table.AddColor("Orange", new ExcelRange(0, 0, 0, columnCount - 1));
            // Phases
            table.AddColor("Blue", new ExcelRange(1, 0, _maxStates, 0));

            table.SetHeaderRange(0, 1, 0, columnCount - 1);

            return table;
        }
        public ExcelTable DecorateDuplicatesTable(object[,] data)
        {
            return new DuplicatesTable(data);
        }
        public ExcelTable DecorateFrequencyTable(object[,] data, bool isTotal)
        {
            ExcelTable table = new ExcelTable(data);

            // Title
            table.AddColor((isTotal ? "Dark" : "") + "Orange", new ExcelRange(0, 0, 0, 0));
            // Header
            table.AddColor((isTotal ? "Dark" : "") + "Blue", new ExcelRange(1, 0, 1, _maxStates * 2 - 1));

            table.SetHeaderRange(1, 0, 1, _maxStates * 2 - 1);

            return table;
        }
        public ExcelTable DecorateCustomFrequencyTable(object[,] data, int numberOfFrequencyRanges, bool isTotal)
        {
            ExcelTable table = isTotal ? new FrequencyTableWithChart(data) : new ExcelTable(data);

            // Title
            table.AddColor((isTotal ? "Dark" : "") + "Orange", new ExcelRange(0, 0, 0, 0));
            // Header
            table.AddColor((isTotal ? "Dark" : "") + "Blue", new ExcelRange(1, 0, 1, _maxStates));
            // Ranges (We add +1 to range number because (>) range gets added automatically) 
            // for example if last range is 20-30, >30 will be added and we have to account for that
            table.AddColor("Gray", new ExcelRange(2, 0, numberOfFrequencyRanges + 1, 0));

            table.SetHeaderRange(1, 1, 1, _maxStates);

            return table;
        }
        public ExcelTable DecorateClusterDataTable(object[,] data, List<TimeStamp> timeStamps, int clusterTime)
        {
            ExcelTable table = new ExcelTable(data);
            TimeStamp cur;
            ExcelRange cRange;
            ExcelResources excelResources = ExcelResources.GetInstance();
            // Go through timestamps and add appropriate coloring
            for (int i = 0; i < timeStamps.Count; i++)
            {
                cRange = new ExcelRange(i, 0, i, 1);
                cur = timeStamps[i];
                // Cluster time and wakefulness - dark red
                if (cur.TimeDifferenceInSeconds >= clusterTime && cur.State == excelResources.MaxStates)
                    table.AddColor("DarkRed", cRange);
                // Wakefulness - red
                else if (cur.State == excelResources.MaxStates)
                    table.AddColor("Red", cRange);
                // Sleep - yellow
                else if (cur.State == excelResources.MaxStates - 1)
                    table.AddColor("Yellow", cRange);
                // PS - green
                else if (cur.State == excelResources.MaxStates - 2)
                    table.AddColor("Green", cRange);
            }

            return table;
        }
    }
}
