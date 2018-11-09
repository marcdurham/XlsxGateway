using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace XlsxGateway.Gateways
{
    public class SharedStringXmlGateway : XmlGateway, ISharedStringGateway
    {
        private const string ValueXPath = @"//default:t";
        private const string ParentElementName = @"si";
        private const string ChildElementName = @"t";
        private const string SaveErrorMessage = @"Error saving shared strings for Excel worksheet to stream: ";

        private XmlDocument document;

        public List<string> SharedStrings;

        public void OpenFrom (string streamText)
        {
            var doc = new XmlDocument ();
            doc.LoadXml (streamText);
            document = doc;
            SharedStrings = ExtractStringsFrom (document);
        }
        
        public void Save (Stream targetSharedStringsStream)
        {
            try {
                document.Save (targetSharedStringsStream);
            } catch (Exception e) {
                throw new ExcelSheetException (SaveErrorMessage + e.Message);
            }
        }

        public void Add (string value)
        {
            SharedStrings.Add (value);

            XmlElement parent = document.CreateElement (
                ParentElementName, 
                DefaultNameSpaceUrl);
            
            XmlElement child = document.CreateElement (
                ChildElementName, 
                DefaultNameSpaceUrl);
            
            child.InnerText = value;
            parent.AppendChild (child);
            document.DocumentElement.AppendChild (parent);

            var sst = (XmlElement)document.GetElementsByTagName("sst")[0];
            sst.SetAttribute("count", SharedStrings.Count.ToString());
        }

        public string StringAtIndexOf (string value)
        {
            return SharedStrings [int.Parse (value)];
        }

        public int Count { get { return SharedStrings.Count; } }

        static List<string> ExtractStringsFrom (XmlDocument document)
        {
            var strings = new List<string> ();

            XmlNodeList stringNodes = document.SelectNodes (
                ValueXPath, 
                NameSpaceManagerFrom (document));

            foreach (XmlNode node in stringNodes)
                strings.Add (node.InnerText);

            return strings;
        }
    }
}

