using DataProcessing.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    internal class TableDecorator
    {
        public TableCollection DecorateRawData(List<DataTable> tables, List<TimeStamp> timeStamps)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasHeader = false;
            collection.HasTiteOnTop = false;

            // Get marked record indexes for coloring
            List<ColorRange> rowIndexes = new List<ColorRange>();
            // Default hour distinction colors
            int time = 0;
            TimeStamp cur;
            for (int i = 0; i < timeStamps.Count; i++)
            {
                cur = timeStamps[i];
                time += cur.TimeDifferenceInSeconds;
                if (time == 3600) { rowIndexes.Add(new ColorRange(0, i, 4, i)); time = 0; }
                if (time > 3600) { throw new Exception("Incorrect time mark division."); }
            }
            collection.ColorRanges.Add("Green", rowIndexes.ToArray());
            rowIndexes.Clear();
            // Indexes where we inserted hour break
            int[] timeMarkedindexes = Enumerable.Range(0, timeStamps.Count).Where(i => timeStamps[i].IsTimeMarked).ToArray();
            // Indexes if file doesn't contain 10 am (and another one) and we had to insert it (colored yellow)
            int[] markerIndexes = Enumerable.Range(0, timeStamps.Count).Where(i => timeStamps[i].IsMarker).ToArray();
            for (int i = 0; i < timeMarkedindexes.Length; i++)
            {
                rowIndexes.Add(new ColorRange(0, timeMarkedindexes[i], 4, timeMarkedindexes[i]));
            }
            collection.ColorRanges.Add("DarkGreen", rowIndexes.ToArray());

            // Create and add ColorRange array for yellow (Markers)
            rowIndexes.Clear();
            for (int i = 0; i < markerIndexes.Length; i++)
            {
                rowIndexes.Add(new ColorRange(0, markerIndexes[i], 0, markerIndexes[i]));
            }
            collection.ColorRanges.Add("Yellow", rowIndexes.ToArray());

            return collection;
        }
        public TableCollection DecorateLatencyTable(List<DataTable> tables)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasHeader = true;
            collection.HasTiteOnTop = true;

            // Header
            collection.ColorRanges.Add("Orange", new ColorRange[] { new ColorRange(0, 0, 0, 0) });
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 1, 1, 1) });

            return collection;

        }
        public TableCollection DecorateStatTables(List<DataTable> tables, int criteriaNumber)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = true;
            collection.HasHeader = true;
            collection.HasTiteOnTop = false;
            collection.AutofitFirstColumn = true;
            // If we don't have any table we set column number to 0 (because in that case we don't want to format anything)
            // otherwise set the column number (Also we subtract 1 because correct range selection in excel relative to position)
            int columnNumber = (tables == null || tables.Count == 0) ? 0 : tables[0].Columns.Count - 1;
            collection.RightAlignmentRange = new ColorRange(1, 0, columnNumber, 0);

            // Header
            collection.ColorRanges.Add("Orange", new ColorRange[] { new ColorRange(0, 0, 4, 0) });
            // Phases
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 1, 0, 3) });
            // Specific criterias
            if (criteriaNumber != 0)
            {
                collection.ColorRanges.Add("Red", new ColorRange[] { new ColorRange(0, 4, 0, 4 + criteriaNumber) });
            }

            return collection;
        }
        public TableCollection DecorateGraphTables(List<DataTable> tables)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasHeader = true;
            collection.HasTiteOnTop = false;
            collection.AutofitFirstColumn = true;
            // If we don't have any table we set column number to 0 (because in that case we don't want to format anything)
            // otherwise set the column number (Also we subtract 1 because correct range selection in excel relative to position)
            int columnNumber = (tables == null || tables.Count == 0) ? 0 : tables[0].Columns.Count - 1;
            collection.RightAlignmentRange = new ColorRange(1, 0, columnNumber, 0);

            int columnCount = tables[0].Columns.Count;
            // Header
            collection.ColorRanges.Add("Orange", new ColorRange[] { new ColorRange(0, 0, columnCount - 1, 0) });
            // Phases
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 1, 0, 3) });

            return collection;
        }
        public TableCollection DecorateDuplicatesTable(List<DataTable> tables)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasHeader = false;
            collection.HasTiteOnTop = false;

            return collection;
        }
        public TableCollection DecorateFrequencyTables(List<DataTable> tables)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = true;
            collection.HasHeader = true;
            collection.HasTiteOnTop = true;

            // Header
            collection.ColorRanges.Add("Orange", new ColorRange[] { new ColorRange(0, 0, 0, 0) });
            // For scalability it would be better to make this dynamic and select range based on max states
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 1, 5, 1) });

            return collection;
        }
        public TableCollection DecorateCustomFrequencyTables(List<DataTable> tables, int numberOfFrequencyRanges)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = true;
            collection.HasHeader = true;
            collection.HasTiteOnTop = true;

            // Header
            collection.ColorRanges.Add("Orange", new ColorRange[] { new ColorRange(0, 0, 0, 0) });
            // For scalability it would be better to make this dynamic and select range based on max states
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 1, 3, 1) });
            // Ranges (We add +1 to range number because (>) range gets added automatically) 
            // for example if last range is 20-30, >30 will be added and we have to account for that
            collection.ColorRanges.Add("Gray", new ColorRange[] { new ColorRange(0, 2, 0, numberOfFrequencyRanges + 1 ) });

            return collection;
        }
        public TableCollection DecorateClusterDataTable(List<DataTable> tables, List<TimeStamp> timeStamps, int clusterTime)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasHeader = false;
            collection.HasTiteOnTop = false;

            List<ColorRange> darkReds = new List<ColorRange>();
            List<ColorRange> reds = new List<ColorRange>();
            List<ColorRange> yellows = new List<ColorRange>();
            List<ColorRange> greens = new List<ColorRange>();

            TimeStamp cur;
            ColorRange cRange;
            // Go through timestamps and add appropriate coloring
            for (int i = 0; i < timeStamps.Count; i++)
            {
                cRange = new ColorRange(0, i, 1, i);
                cur = timeStamps[i];
                // Cluster time and wakefulness - dark red
                if (cur.TimeDifferenceInSeconds >= clusterTime && cur.State == 3)
                    darkReds.Add(cRange);
                // Wakefulness - red
                else if (cur.State == 3)
                    reds.Add(cRange);
                // Sleep - yellow
                else if (cur.State == 2)
                    yellows.Add(cRange);
                // PS - green
                else if (cur.State == 1)
                    greens.Add(cRange);
            }

            collection.ColorRanges.Add("DarkRed", darkReds.ToArray());
            collection.ColorRanges.Add("Red", reds.ToArray());
            collection.ColorRanges.Add("Yellow", yellows.ToArray());
            collection.ColorRanges.Add("Green", greens.ToArray());

            return collection;
        }
    }
}
