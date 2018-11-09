using System.Collections.Generic;
using XlsxGateway.Gateways;
using XlsxGateway.Models;

namespace XlsxGateway.UnitTests
{
    public class WorksheetGatewayFake : IWorksheetGateway
    {
        public void Close()
        {
        }

        public Worksheet OpenFromReturns;
        public Worksheet OpenFrom(string fileName, string sheetName)
        {
            return OpenFromReturns;
        }

        public void SaveAndClose()
        {
        }

        public Worksheet SaveToSheet;
        public void SaveTo(Worksheet sheet)
        {
            SaveToSheet = sheet;
        }

        public void UpdateWith(
            string sourceFileName, 
            string targetFileName, 
            string sheetName, 
            Worksheet updatedWorksheet)
        {
        }

        public List<string> SheetNamesFrom(string fileName)
        {
            return null;
        }
    }
}
