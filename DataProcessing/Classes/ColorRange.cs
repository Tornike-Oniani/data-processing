using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataProcessing.Classes
{
    /// <summary>
    /// Range of a table that we use to color it in excel
    /// </summary>
    internal class ColorRange
    {
        public int StartColumn { get; private set; }
        public int StartRow { get; private set; }
        public int EndColumn { get; private set; }
        public int EndRow { get; private set; }

        public ColorRange(int startColumn, int startRow, int endColumn, int endRow)
        {
            this.StartColumn = startColumn;
            this.StartRow = startRow;
            this.EndColumn = endColumn;
            this.EndRow = endRow;
        }
    }
}
