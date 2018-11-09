using System;
using System.Collections.Generic;
using System.Xml;
using XlsxGateway.Models;

namespace XlsxGateway.Gateways
{
    public class SheetDocumentXmlReader : XmlGateway, ISheetDocumentReader
    {
        private const string RowXPath = @"//default:row";
        private const string ReferenceAttribute = @"r";
        private const string CellTypeAttribute = @"t";
        private const string CellStyleAttribute = @"s";
        private const string NumberType = @"n";
        private const string StringType = @"s";
        private const string InlineStringType = @"inlineStr";
        private const string NoRowsPresent = @"No rows are present!";

        private ISharedStringGateway sharedStringGateway;

        public SheetDocumentXmlReader (ISharedStringGateway sharedStringGateway)
        {
            this.sharedStringGateway = sharedStringGateway;
        }

        public Worksheet OpenFrom (XmlDocument xmlDocument)
        {
            return new RowWorksheet(RowsFrom(xmlDocument));
        }

        public Cell CellFrom (XmlNode cellNode)
        {
            if (cellNode == null)
                return Cell.Empty;

            CellType type = CellTypeFrom(cellNode);
            string value = cellNode.InnerText;

           if (type == CellType.SharedString)
                value = sharedStringGateway.StringAtIndexOf(value);

            return new Cell () {
                Row = RowNumberFromCell (cellNode),
                Column = ColumnNameFrom (cellNode),
                Value = value,
                Type = type,
                Style = StyleFrom(cellNode)
            };
        }

        List<Row> RowsFrom (XmlDocument document)
        {
            var rows = new List<Row> ();

            foreach (XmlNode rowNode in RowNodesFrom (document))
                rows.Add (RowFrom (rowNode));

            return rows;
        }

        Row RowFrom (XmlNode rowNode)
        {
            return new Row () {
                RowNumber = RowNumberFrom (rowNode),
                Cells = CellsFrom (rowNode)
            };
        }

        Dictionary<string, Cell> CellsFrom (XmlNode rowNode)
        {
            var cells = new Dictionary<string, Cell> ();

            foreach (XmlNode cellNode in CellNodesFrom (rowNode)) {
                var cell = CellFrom (cellNode);
                cells.Add (cell.Column, cell);
            }

            return cells;
        }

        static int RowNumberFrom (XmlNode rowNode)
        {
            return int.Parse (rowNode.Attributes [ReferenceAttribute].Value);
        }

        static XmlNodeList CellNodesFrom (XmlNode rowNode)
        {
            return rowNode.ChildNodes;
        }

        static XmlNodeList RowNodesFrom (XmlDocument document)
        {
            XmlNodeList rowNodes = document.SelectNodes (
                RowXPath, 
                NameSpaceManagerFrom (document));

            if (rowNodes.Count == 0)
                throw new ExcelSheetException (NoRowsPresent);

            return rowNodes;
        }

        static CellType CellTypeFrom (XmlNode cellNode)
        {
            try
            {
                var cellElement = cellNode as XmlElement;
                string type = cellElement.GetAttribute(CellTypeAttribute);

                if (type == null)
                    return CellType.Number;
                else if (string.Equals(type, NumberType, StringComparison.OrdinalIgnoreCase))
                    return CellType.Number;
                else if (string.Equals(type, StringType, StringComparison.OrdinalIgnoreCase))
                    return CellType.SharedString;
                else if (string.Equals(type,InlineStringType, StringComparison.OrdinalIgnoreCase))
                    return CellType.InlineString;

                // In .xlsx files if the XML type attribute "t" is missing
                // then the type is the number "n" type.
                return CellType.Number;
            }
            catch(Exception)
            {
                throw;
            }
        }

        static string ColumnNameFrom (XmlNode node)
        {
            return CellAddress.From (node.Attributes [ReferenceAttribute].Value)
                .Column;
        }

        static int RowNumberFromCell (XmlNode node)
        {
            return CellAddress.From (node.Attributes [ReferenceAttribute].Value)
                .Row;
        }

        static int? StyleFrom(XmlNode node)
        {
            var element = node as XmlElement;
            string style = element.GetAttribute(CellStyleAttribute);

            if (string.IsNullOrWhiteSpace(style))
                return null;

            return int.Parse(style);
        }
    }
}

