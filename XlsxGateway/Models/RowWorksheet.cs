using System.Collections.Generic;
using System.Linq;

namespace XlsxGateway.Models
{
    public class RowWorksheet : Worksheet
    {
        public Dictionary<string, string> headerColumns;
        public List<string> headers;
        public List<Row> rows;

        public RowWorksheet (IEnumerable<Row> rows)
        {
            this.rows = rows.ToList ();
            headerColumns = ExtractHeaderColumnsFrom (rows);
        }

        public override IEnumerable<Row> Rows ()
        {
            return rows;
        }

        public override Row Row (int rowIndex)
        {
            return rows [rowIndex];
        }

        public override List<string> RowsFromKeyColumn (string columnName)
        {
            return RowsFromColumn (columnName, rowCount: int.MaxValue, stopAtNull: true);
        }

        public override List<string> RowsFromColumn (string columnName, int rowCount)
        {
            return RowsFromColumn (columnName, rowCount, stopAtNull: false);
        }

        public override void SetCellValueAt (string cellAddress, string value)
        {
            string column = CellAddress.From (cellAddress).Column;
            int rowNumber = CellAddress.From (cellAddress).Row;

            SetCellValueAtColumn (column, rowNumber, value, CellType.SharedString);
        }

        public override void SetCellValueAt(string cellAddress, decimal value)
        {
            string column = CellAddress.From(cellAddress).Column;
            int rowNumber = CellAddress.From(cellAddress).Row;

            SetCellValueAtColumn(column, rowNumber, value.ToString(), CellType.Number);
        }

        public override void SetCellValue (string columnName, int rowIndex, string value)
        {
            string column = string.Empty;

            if (!headerColumns.TryGetValue(columnName, out column))
                throw new ExcelSheetException("Can not find a column header named: " + columnName);

            SetCellValueAtColumn (column, rowIndex + 1, value, CellType.SharedString);
        }

        public override void SetCellValue(string columnName, int rowIndex, decimal value)
        {
            string column = string.Empty;

            if (!headerColumns.TryGetValue(columnName, out column))
                throw new ExcelSheetException("Can not find a column header named: " + columnName);

            SetCellValueAtColumn(column, rowIndex + 1, value.ToString(), CellType.Number);
        }

        public override bool Contains (Row row)
        {
            return RowExistsAt (row.RowNumber);
        }

        bool RowExistsAt (int rowNumber)
        {
            return rows.Exists (r => r.RowNumber == rowNumber);
        }

        private void SetCellValueAtColumn (string column, int rowNumber, string value, CellType type)
        {
            bool valueIsNull = value == null;

            Row row = RowAt (rowNumber);
            if (row == null && !valueIsNull) {
                row = new Row { RowNumber = rowNumber };
                rows.Add (row);
            }

            Cell cell;

            if (!row.Cells.TryGetValue (column, out cell)) {
                if (!valueIsNull)
                    row.AddCell (column, value, type);
            } else {
                if (!valueIsNull)
                    cell.Value = value;
                else {
                    row.Cells.Remove (column);
                    if (row.Cells.Count == 0)
                        rows.Remove (row);
                }
            }
        }

        public override bool CellExistsAt (string cellAddress)
        {
            Row row = RowAt (CellAddress.From (cellAddress).Row);
            if (row == null)
                return false;

            Cell cell;
            if (!row.Cells.TryGetValue (CellAddress.From (cellAddress).Column, out cell))
                return false;

            return true;
        }

        public override string GetCellValueAt (string cellAddress)
        {
            Cell cell = GetCellAt(cellAddress);

            return cell == null ? null : cell.Value;
        }

        public override Cell GetCellAt(string cellAddress)
        {
            string column = CellAddress.From(cellAddress).Column;
            int rowNumber = CellAddress.From(cellAddress).Row;

            return GetCell(column, rowNumber - 1) ?? Cell.Empty;
        }

        public override Row AddRow (int rowNumber)
        {
            if (rows.Count (r => r.RowNumber == rowNumber) > 0)
                rows.Remove (rows.Single (r => r.RowNumber == rowNumber));

            var row = new Row () { RowNumber = rowNumber };
            rows.Add (row);

            return row;
        }

        public override int RowCount ()
        {
            return rows.Count;
        }

        public override bool Contains (Cell cell)
        {
            return CellExistsAt (cell.Address);
        }

        private Cell GetCell(string columnName, int rowIndex)
        {
            Row row = RowAt(rowIndex + 1);
            if (row == null)
                return null;

            Cell cell = new Cell();
            if (!row.Cells.TryGetValue(columnName, out cell))
                return null;

            return cell;
        }

        private Row RowAt (int rowNumber)
        {
            return rows.SingleOrDefault (r => r.RowNumber == rowNumber);
        }

        private List<string> RowsFromColumn (string columnName, int rowCount, bool stopAtNull)
        {
            var rowValues = new List<string> ();

            string column = string.Empty;
            if (!headerColumns.TryGetValue (columnName, out column))
                throw new ExcelSheetException ("ColumnName: " + columnName + " not found in headers.");

            foreach (var row in rows.Where (r => r.RowNumber > 1)) {
                Cell cell;
                if (row.Cells.TryGetValue (column, out cell)) {
                    rowValues.Add (cell.Value);
                } else {
                    if (stopAtNull)
                        break;
                    rowValues.Add (null);
                }
            }

            if (!stopAtNull && rowValues.Count < rowCount) {
                int remaining = rowCount - rowValues.Count;
                for (int i = 0; i < remaining; i++)
                    rowValues.Add (null);
            }

            return rowValues;
        }

        private Dictionary<string, string> ExtractHeaderColumnsFrom (IEnumerable<Row> rows)
        {
            var headerCells = new Dictionary<string, string> ();

            var headerRow = rows.Where(r => r.RowNumber == 1).SingleOrDefault();

            if (headerRow == null)
                throw new ExcelSheetException("Missing Header Row 1");

            foreach (var keyValue in headerRow.Cells)
            {
                if (headerCells.ContainsKey(keyValue.Value.Value))
                    throw new ExcelSheetException("Header already contains column named: " + keyValue.Value.Value);
                headerCells.Add(keyValue.Value.Value, keyValue.Value.Column);
            }

            return headerCells;
        }
    }
}