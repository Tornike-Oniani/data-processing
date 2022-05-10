using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    internal class TableCollection
    {
        // Tables ranging one once column having the same style
        public List<DataTable> Tables { get; set; }
        // Does Tables list contain total stat on top
        public bool HasTotal { get; set; }
        /// <summary>
        /// If the tables title and column names should be displayed in excel
        /// Some tabels don't need title or column in excel file (for example raw data or cluster data)
        /// </summary>
        public bool HasHeader { get; set; }
        // Some tables have title on top of column names instead of the same row
        public bool HasTiteOnTop { get; set; }
        /// <summary>
        /// Which rows should be colored (string will be a name of a color e. g. "blue" and then we will have string color names mapped to actual Color class)
        /// Class ColorRange contains start and end of columns and rows which should be colored by the key
        /// </summary>
        public Dictionary<string, ColorRange[]> ColorRanges { get; set; }

        // Constructor
        public TableCollection()
        {
            // Init
            ColorRanges = new Dictionary<string, ColorRange[]>();
        }
    }
}
