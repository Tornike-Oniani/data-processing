using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace DataProcessing.Classes.Export
{
    internal class ExcelTable : IExportable
    {
        // Table data that should be written on excel file
        protected object[,] _data;
        // List of all colors and ranges that will be used for decorating this table
        // in excel file. We keep here name of the color as key and value all the ranges
        // that should be colored with that color. Here the color is string but in excel
        // manager we will map it to actual Color class
        public Dictionary<string, List<ExcelRange>> ColorRanges { get; private set; }

        // Constructor
        public ExcelTable(object[,] data)
        {
            this._data = data;
            ColorRanges = new Dictionary<string, List<ExcelRange>>();
        }

        public void AddColor(string colorName, ExcelRange range)
        {
            // If we have no range with color initialize it first
            if (!ColorRanges.ContainsKey(colorName))
            {
                ColorRanges.Add(colorName, new List<ExcelRange>() { range });
                return;
            }

            ColorRanges[colorName].Add(range);
        }

        public virtual void ExportToSheet(_Worksheet sheet, int verticalPosition, int horizontalPosition)
        {
            throw new System.NotImplementedException();
        }
    }
}
