using XlsxGateway.Models;
using System.Xml;

namespace XlsxGateway.Gateways
{
    public interface ISheetDocumentSaver
    {
        XmlDocument Update(XmlDocument document, Worksheet from);
        XmlDocument ResetStringTypes(XmlDocument document, WorksheetStringType toStringType);
        Worksheet OpenFrom(XmlDocument document);
    }
}
