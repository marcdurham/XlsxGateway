using System.Xml;

namespace XlsxGateway.Gateways
{
    public abstract class XmlGateway
    {
        protected const string DefaultNameSpaceUrl = @"http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        protected const string DefaultNameSpace = @"default";

        public static XmlNamespaceManager NameSpaceManagerFrom (XmlDocument document)
        {
            var nameSpaceManager = new XmlNamespaceManager (document.NameTable);
            nameSpaceManager.AddNamespace (DefaultNameSpace, DefaultNameSpaceUrl);
            return nameSpaceManager;
        }
    }
}

