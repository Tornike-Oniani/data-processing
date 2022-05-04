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
            collection.HasTitle = false;

            // Get marked record indexes for coloring
            List<ColorRange> rowIndexes = new List<ColorRange>();
            // Indexes where we inserted hour break
            int[] timeMarkedindexes = Enumerable.Range(0, timeStamps.Count).Where(i => timeStamps[i].IsTimeMarked).ToArray();
            // Indexes if file doesn't contain 10 am (and another one) and we had to insert it (colored yellow)
            int[] markerIndexes = Enumerable.Range(0, timeStamps.Count).Where(i => timeStamps[i].IsMarker).ToArray();
            for (int i = 0; i < timeMarkedindexes.Length; i++)
            {
                rowIndexes.Add(new ColorRange(0, timeMarkedindexes[i], 0, timeMarkedindexes[i]));
            }
            collection.ColorRanges.Add("Green", rowIndexes.ToArray());

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
            collection.HasTitle = true;
            // Header
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 0, 1, 0) });

            return collection;

        }
        public TableCollection DecorateStatTables(List<DataTable> tables, int criteriaNumber)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = true;
            collection.HasTitle = false;
            // Header
            collection.ColorRanges.Add("Orange", new ColorRange[] { new ColorRange(0, 0, 4, 0) });
            // Phases
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 1, 0, 3) });
            // Specific criterias
            collection.ColorRanges.Add("Red", new ColorRange[] { new ColorRange(0, 4, 0, 4 + criteriaNumber) });

            return collection;
        }
        public TableCollection DecorateGraphTables(List<DataTable> tables)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasTitle = false;
            int columnCount = tables[0].Columns.Count;
            // Header
            collection.ColorRanges.Add("Orange", new ColorRange[] { new ColorRange(0, 0, columnCount, 0) });
            // Phases
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 1, 0, 3) });

            return collection;
        }
        public TableCollection DecorateDuplicatesTable(List<DataTable> tables)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasTitle = false;

            return collection;
        }
        public TableCollection DecorateFrequencyTables(List<DataTable> tables)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = true;
            collection.HasTitle = true;
            // Header
            // For scalability it would be better to make this dynamic and select range based on max states
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 0, 5, 0) });

            return collection;
        }
        public TableCollection DecorateCustomFrequencyTables(List<DataTable> tables, int numberOfFrequencyRanges)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = true;
            collection.HasTitle = true;
            // Header
            // For scalability it would be better to make this dynamic and select range based on max states
            collection.ColorRanges.Add("Blue", new ColorRange[] { new ColorRange(0, 0, 3, 0) });
            // Ranges
            collection.ColorRanges.Add("Gray", new ColorRange[] { new ColorRange(0, 1, 0, numberOfFrequencyRanges) });

            return collection;
        }
        public TableCollection DecorateClusterDataTable(List<DataTable> tables, List<TimeStamp> timeStamps, int clusterTime)
        {
            TableCollection collection = new TableCollection();
            collection.Tables = tables;
            collection.HasTotal = false;
            collection.HasTitle = false;

            List<ColorRange> darkReds = new List<ColorRange>();
            List<ColorRange> reds = new List<ColorRange>();
            List<ColorRange> yellows = new List<ColorRange>();
            List<ColorRange> greens = new List<ColorRange>();

            TimeStamp cur;
            ColorRange cRange;
            // Go through timestamps and add appropriate coloring
            for (int i = 0; i < timeStamps.Count; i++)
            {
                cRange = new ColorRange(0, i, 0, i);
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

            collection.ColorRanges.Add("Dark red", darkReds.ToArray());
            collection.ColorRanges.Add("Red", reds.ToArray());
            collection.ColorRanges.Add("Yellow", yellows.ToArray());
            collection.ColorRanges.Add("Green", greens.ToArray());

            return collection;
        }
    }
}
