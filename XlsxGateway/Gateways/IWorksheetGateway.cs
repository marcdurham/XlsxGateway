using System.Collections.Generic;
using XlsxGateway.Models;

namespace XlsxGateway.Gateways
{
    public interface IWorksheetGateway
    {
        Worksheet OpenFrom (string fileName, string sheetName);  
        void SaveTo (Worksheet sheet);
        // Five layers: file name + sheet name -> zip archive + entry + Streams -> XmlDocuments + XmlNodes -> Worksheets + Cells
        List<string> SheetNamesFrom(string fileName);
    }
}

