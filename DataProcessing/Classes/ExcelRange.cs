namespace DataProcessing.Classes
{
    /// <summary>
    /// Range of a table that we use to color it in excel
    /// </summary>
    internal class ExcelRange
    {
        public int StartColumn { get; private set; }
        public int StartRow { get; private set; }
        public int EndColumn { get; private set; }
        public int EndRow { get; private set; }

        public ExcelRange(int startRow, int startColumn, int endRow, int endColumn)
        {
            this.StartRow = startRow;
            this.StartColumn = startColumn;
            this.EndRow = endRow;
            this.EndColumn = endColumn;
        }
    }
}
