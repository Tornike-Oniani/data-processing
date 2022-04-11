using DataProcessing.Utils;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    /// <summary>
    /// Each TableCollection corresponds to one sheet in excel file, sometimes we can have multiple different kinds of tables on one sheet
    /// that's why whe have list inside of list in tables. Each separate list of tables will be apart from each other by 2 excel columns
    /// horizontally
    /// </summary>
    internal class TableCollection
    {
        // Name of the excel sheet where this tables will be exported
        public string Name { get; set; }
        // Tables containing processed data
        public List<List<DataTableInfo>> Tables { get; set; }
        // Excel range (start row, start column, end row, end column) for coloring header
        public int[] HeaderIndexes { get; set; }
        // Excel range (start row, start column, end row, end column) for coloring phases
        public int[] PhaseIndexes { get; set; }
        // Excel range (start row, start column, end row, end column) for coloring criterias (only stats table has this)
        public int[] CriteriaIndexes { get; set; }
        // Excel range (start row, start column, end row, end colum) for appending it to chart
        public int[] ChartDataIndexes { get; set; }
    }
}
