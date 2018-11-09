using System.Xml;
using XlsxGateway.Models;

namespace XlsxGateway.Gateways
{
    public interface ISheetDocumentReader
    {
        Worksheet OpenFrom (XmlDocument xmlDocument); 
        Cell CellFrom(XmlNode node);
    }
}