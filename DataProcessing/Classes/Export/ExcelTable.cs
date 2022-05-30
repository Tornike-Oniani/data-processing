using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Drawing;

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
            this.ColorRanges = new Dictionary<string, List<ExcelRange>>();
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

        // IExportable interface
        public virtual int ExportToSheet(_Worksheet sheet, int verticalPosition, int horizontalPosition)
        {
            WriteData(sheet, verticalPosition, horizontalPosition);
            Decorate(sheet, verticalPosition, horizontalPosition);
            return verticalPosition + _data.GetLength(1);
        }

        // Main export functions
        protected void WriteData(_Worksheet sheet, int verticalPosition, int horizontalPosition)
        {
            int rowCount = _data.GetLength(0) - 1;
            int colCount = _data.GetLength(1) - 1;
            Range range = GetRange(sheet, verticalPosition, horizontalPosition, verticalPosition + rowCount, horizontalPosition + colCount);
            range.Value = _data;
        }
        protected void Decorate(_Worksheet sheet, int verticalPosition, int horizontalPosition)
        {
            string colorName;
            List<ExcelRange> ranges;
            Color color;
            Range excRange;
            int startRow;
            int startColumn;
            int endRow;
            int endColumn;

            foreach (KeyValuePair<string, List<ExcelRange>> entry in ColorRanges)
            {
                colorName = entry.Key;
                ranges = entry.Value;
                // Get appropriate color from dictionary
                color = ExcelResources.GetInstance().Colors[colorName];

                foreach (ExcelRange range in ranges)
                {
                    // Set relative positions (ColorRange keeps track of ranges relative to table disregarding current position on excel)
                    startRow = range.StartRow + verticalPosition;
                    startColumn = range.StartColumn + horizontalPosition;
                    endRow = range.EndRow + verticalPosition;
                    endColumn = range.EndColumn + horizontalPosition;

                    // Get range and set its color
                    excRange = GetRange(sheet, startRow, startColumn, endRow, endColumn);
                    excRange.Interior.Color = color;
                }
            }
        }

        // Helper functions
        protected Range GetRange(_Worksheet sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            Range start = sheet.Cells[startRow, startColumn];
            Range end = sheet.Cells[endRow, endColumn];
            return sheet.Range[start, end];
        }
    }
}
