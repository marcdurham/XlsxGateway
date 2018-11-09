using System.Collections.Generic;

namespace XlsxGateway.Gateways
{
    public interface ISheetNameIdGateway
    {
        void OpenFrom (string streamText);
        Dictionary<string, int> ExtractSheetNames ();
        int SheetIdFrom (string sheetName, string fileName);
    }
}

