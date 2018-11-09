using System.Collections.Generic;
using System.Xml;

namespace XlsxGateway.Gateways
{
    public class SheetNameIdXmlGateway : XmlGateway, ISheetNameIdGateway
    {
        private const string SheetNameXPath = @"//default:sheet";
        private const string SheetNameAttributeName = @"sheetName";
        private const string SheetNameSheetIdAttributeName = @"sheetId";
        private const string SheetNameNameAttributeName = @"name";
        private const string SheetNameDoesNotExistMessage = @"The sheet name '{0}' does not exist in the workbook {1}.";

        public Dictionary<string, int> SheetIds;

        XmlDocument workbookDocument;

        public void OpenFrom (string streamText)
        {
            var doc = new XmlDocument ();
            doc.LoadXml (streamText);
            workbookDocument = doc;
            SheetIds = ExtractSheetNames ();
        }

        public Dictionary<string, int> ExtractSheetNames ()
        {
            return ExtractFrom (workbookDocument);
        }

        public int SheetIdFrom (string sheetName, string fileName)
        {
            int sheetId = 0;
            SheetIds.TryGetValue (sheetName, out sheetId);

            if (sheetId == 0)
                throw new ExcelSheetException (string.Format (SheetNameDoesNotExistMessage, sheetName, fileName));

            return sheetId;
        }

        static Dictionary<string, int> ExtractFrom (XmlDocument document)
        {
            var names = new Dictionary<string, int> ();

            XmlNodeList nameNodes = document.SelectNodes (
                SheetNameXPath,
                NameSpaceManagerFrom (document));

            foreach (XmlNode node in nameNodes) {
                XmlElement element = node as XmlElement;
                int sheetId = int.Parse (element.GetAttribute (SheetNameSheetIdAttributeName));
                string name = element.GetAttribute (SheetNameNameAttributeName);
                names.Add (name, sheetId);
            }

            return names;
        }
    }
}

