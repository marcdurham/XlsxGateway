using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using XlsxGateway.Models;

namespace XlsxGateway.Gateways
{
    public class SheetDocumentXmlSaver : XmlGateway, ISheetDocumentSaver
    {
        public List<Row> rows;

        private const string SheetDataXPath = @"//default:sheetData";
        private const string ReferenceAttributeName = @"r";
        private const string CellTypeAttributeName = @"t";
        private const string CellStyleAttributeName = @"s";
        private const string NumberType = @"n";
        private const string SharedStringType = @"s";
        private const string InlineStringType = @"inlineStr";
        private const string CellXQueryByAddress = @"//default:c[@r='{0}']";
        private const string RowXQueryByNumber = @"//default:row[@r='{0}']";
        private const string RowElementName = @"row";
        private const string CellElementName = @"c";
        private const string CellValueElementName = @"v";
        private const string UpdateErrorMessage = @"Error Updating Excel XML Worksheet in stream: ";
        private const string SaveErrorMessage = @"Error saving worksheet to stream: ";

        private ISharedStringGateway sharedStringGateway;
        private ISheetDocumentReader sheetDocumentReader;
        private XmlDocument targetDocument;
        private Worksheet targetSheet;
        private WorksheetStringType targetStringType;

        public SheetDocumentXmlSaver(
            ISharedStringGateway sharedStringGateway,
            ISheetDocumentReader sheetDocumentReader)
        {
            this.sharedStringGateway = sharedStringGateway;
            this.sheetDocumentReader = sheetDocumentReader;
        }

        public XmlDocument Update(XmlDocument document, Worksheet from)
        {
            try
            {
                targetDocument = document;

                targetSheet = sheetDocumentReader.OpenFrom(targetDocument);
                targetSheet.StringType = from.StringType;

                DeleteCellsMissingFrom(from);

                AddNewRowsFrom(from.Rows());

                return targetDocument;
            }
            catch (Exception e)
            {
                throw new ExcelSheetException(UpdateErrorMessage + e.Message);
            }
        }

        public XmlDocument ResetStringTypes(XmlDocument document, WorksheetStringType toStringType)
        {
            try
            {
                targetDocument = document;
                targetStringType = toStringType;

                ResetStringTypesIn(targetDocument);

                return targetDocument;
            }
            catch (Exception e)
            {
                throw new ExcelSheetException(UpdateErrorMessage + e.Message);
            }
        }

        private void ResetStringTypesIn(XmlDocument document)
        {
            var cellNodes = document.GetElementsByTagName(CellElementName);

            foreach(XmlElement cellNode in cellNodes)
            {
                var cell = sheetDocumentReader.CellFrom(cellNode);

                SetCellProperties(cellNode, cell);
            }
        }

        public Worksheet OpenFrom(XmlDocument document)
        {
            return sheetDocumentReader.OpenFrom(document);
        }

        // TODO: Finder gateway?
        XmlNode CellNodeForCellAddress(string cellAddress)
        {
            return targetDocument.SelectSingleNode(
                string.Format(CellXQueryByAddress, cellAddress),
                NameSpaceManagerFrom(targetDocument));
        }

        void SetCellProperties(XmlNode cellNode, Cell cell)
        {
            var cellElement = cellNode as XmlElement;

            SetCellValue(cellElement, cell);
            SetCellReference(cellElement, cell);
            SetCellType(cellElement, cell);
            SetCellStyle(cellElement, cell);
        }

        void Delete(Row row)
        {
            XmlNode rowNode = RowNodeForRowNumber(row.RowNumber);
            XmlNode sheetDataNode = targetDocument.SelectSingleNode(
                SheetDataXPath,
                NameSpaceManagerFrom(targetDocument));
            sheetDataNode.RemoveChild(rowNode);
        }

        XmlNode RowNodeForRowNumber(int rowNumber)
        {
            return targetDocument.SelectSingleNode(
                string.Format(RowXQueryByNumber, rowNumber),
                NameSpaceManagerFrom(targetDocument));
        }

        void AddCellNodeFrom(Cell cell)
        {
            // Cells are not added in order
            XmlElement cellElement = targetDocument.CreateElement(
                CellElementName,
                DefaultNameSpaceUrl);

            XmlNode rowNode = RowNodeFrom(cell.Row);
            rowNode.AppendChild(cellElement);

            SetCellProperties(cellElement, cell);
        }

        void AddNewRowsFrom(IEnumerable<Row> sourceRows)
        {
            foreach (Row row in sourceRows)
                AddNewCellsFrom(row.Cells.Select(c => c.Value));
        }

        void SetCellValue(XmlElement cellElement, Cell cell)
        {
            if (IsStringType(cell) && targetStringType == WorksheetStringType.SharedString) 
            {
                sharedStringGateway.Add(cell.Value);

                XmlElement valueElement = targetDocument.CreateElement(
                   CellValueElementName,
                   DefaultNameSpaceUrl);

                valueElement.InnerText = IndexOfStringJustAdded();

                cellElement.RemoveAll();
                cellElement.AppendChild(valueElement);
            }
            else if (IsStringType(cell) 
                && targetStringType == WorksheetStringType.InlineString)
            {
                var si = targetDocument.CreateElement(
                    "is",
                    DefaultNameSpaceUrl);

                var t = targetDocument.CreateElement(
                    "t",
                    DefaultNameSpaceUrl);

                cellElement.RemoveAll();
                cellElement.AppendChild(si);
                si.AppendChild(t);
                t.InnerText = cell.Value;
            }
            else
            {
                XmlElement valueElement = targetDocument.CreateElement(
                    CellValueElementName,
                    DefaultNameSpaceUrl);

                valueElement.InnerText = cell.Value;

                cellElement.RemoveAll();
                cellElement.AppendChild(valueElement);
            }
        }

        bool IsStringType(Cell cell)
        {
            return cell.Type == CellType.SharedString
                || cell.Type == CellType.InlineString;
        }

        XmlNode RowNodeFrom(int rowNumber)
        {
            XmlNode rowNode = RowNodeForRowNumber(rowNumber);

            if (rowNode == null)
                rowNode = CreateRowNodeFrom(rowNumber);

            return rowNode;
        }

        XmlNode CreateRowNodeFrom(int rowNumber)
        {
            XmlNode rowNode;
            XmlNode sheetDataNode = targetDocument.SelectSingleNode(
                SheetDataXPath,
                NameSpaceManagerFrom(targetDocument));

            XmlElement rowElement = targetDocument.CreateElement(
                RowElementName,
                DefaultNameSpaceUrl);

            rowElement.SetAttribute(
                ReferenceAttributeName,
                rowNumber.ToString());

            sheetDataNode.AppendChild(rowElement);
            rowNode = rowElement;

            return rowNode;
        }

        void SetCellType(XmlElement cellElement, Cell cell)
        {
            if (IsStringType(cell)
                && targetStringType == WorksheetStringType.SharedString)
                cellElement.SetAttribute(
                    CellTypeAttributeName,
                    SharedStringType);
            else if (IsStringType(cell)
                && targetStringType == WorksheetStringType.InlineString)
                cellElement.SetAttribute(
                    CellTypeAttributeName,
                    InlineStringType);
            else
                cellElement.RemoveAttribute(CellTypeAttributeName);
        }

        static void SetCellStyle(XmlElement cellElement, Cell cell)
        {
            if (cell.Style == null)
                cellElement.RemoveAttribute(CellStyleAttributeName);
            else
                cellElement.SetAttribute(CellStyleAttributeName, cell.Style.ToString());
        }

        static void SetCellReference(XmlElement cellElement, Cell cell)
        {
            cellElement.SetAttribute(ReferenceAttributeName, cell.Address);
        }

        void DeleteCellsMissingFrom(Worksheet sheet, Row inRow)
        {
            foreach (Cell cell in inRow.Cells.Select(c => c.Value))
            {
                if (!sheet.Contains(cell) && CellIsPresent(cell))
                    Delete(cell);
            }
        }

        void DeleteCellsMissingFrom(Worksheet sourceSheet)
        {
            foreach (Row targetRow in targetSheet.Rows()) 
            {
                if (!sourceSheet.Contains(targetRow))
                    Delete(targetRow);
                else
                    DeleteCellsMissingFrom(sourceSheet, inRow: targetRow);
            }
        }

        bool CellIsPresent(Cell cell)
        {
            XmlNode node = CellNodeForCellAddress(cell.Address);

            return node != null;
        }

        void Delete(Cell targetCell)
        {
            XmlNode rowNode = RowNodeForRowNumber(targetCell.Row);
            XmlNode cellNodeToDelete = CellNodeForCellAddress(
                targetCell.Address);

            rowNode.RemoveChild(cellNodeToDelete);
        }

        void AddNewCellsFrom(IEnumerable<Cell> cells)
        {
            foreach (Cell cell in cells)
            {
                XmlNode targetCellNode = CellNodeForCellAddress(cell.Address);

                if (targetCellNode != null)
                    SyncCellValue(targetCellNode, cell);
                else
                    AddCellNodeFrom(cell);
            }
        }

        void SyncCellValue(XmlNode node, Cell cell)
        {
            Cell cellFromNode = sheetDocumentReader.CellFrom(node);

            if (cellFromNode.Value != cell.Value)
                SetCellProperties(node, cell);
        }

        string IndexOfStringJustAdded()
        {
            return (sharedStringGateway.Count - 1).ToString();
        }

        static bool CellTypeAttributeIsNotNeededFor(Cell cell)
        {
            return cell.Type == CellType.Number;
        }
    }
}
