using Microsoft.Office.Interop.Excel;

namespace DataProcessing.Classes.Export
{
    internal interface IExportable
    {
        int ExportToSheet(_Worksheet sheet, int verticalPosition, int horizontalPosition);
    }
}
