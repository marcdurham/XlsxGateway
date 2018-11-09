using System;
using System.Collections.Generic;
using XlsxGateway.SystemAdapters;
using XlsxGateway.Gateways;

namespace XlsxGateway.Models
{
    public abstract class Worksheet
    {
        public string FileName;
        public string SheetName;
        public WorksheetStringType StringType = WorksheetStringType.InlineString;

        public static Worksheet Open (string sourceFile, string sheetName)
        {
            try
            {
                var sharedStringGatway = new SharedStringXmlGateway ();
                return new ExcelXmlWorksheetGateway(
                    new SystemCompressor(),
                    sharedStringGatway,
                    new SheetNameIdXmlGateway(),
                    new SheetDocumentXmlSaver (
                        sharedStringGatway,
                        new SheetDocumentXmlReader(sharedStringGatway)))
                    .OpenFrom(sourceFile, sheetName);
            }
            catch (Exception e)
            {
                throw new ExcelSheetException (
                    "Error opening Worksheet:"
                    + " sourceFile: " + sourceFile
                    + " sheetName: " + sheetName
                    + " Message: " + e.Message);
            }
        }

        public abstract IEnumerable<Row> Rows ();

        public abstract Row Row (int rowIndex);

        public abstract List<string> RowsFromColumn (string columnName, int rows);

        public abstract List<string> RowsFromKeyColumn (string columnName);

        public abstract bool Contains (Row row);

        public abstract bool CellExistsAt (string cellAddress);

        public abstract bool Contains (Cell cell);

        public abstract void SetCellValue (string columnName, int rowIndex, string value);

        public abstract void SetCellValue(string columnName, int rowIndex, decimal value);

        public abstract void SetCellValueAt (string cellAddress, string value);

        public abstract void SetCellValueAt (string cellAddress, decimal value);

        public abstract string GetCellValueAt (string cellAddress);

        public abstract Cell GetCellAt(string cellAddress);

        public abstract Row AddRow (int rowNumber);

        public abstract int RowCount ();

   }
}
