using System.Collections.Generic;

namespace XlsxGateway.Models
{
    public class Row
    {
        public Dictionary<string, Cell> Cells;
        public int RowNumber;

        public Row ()
        {
            Cells = new Dictionary<string, Cell> ();
        }

        public override string ToString ()
        {
            return string.Format ("Row:{0} Cells:{1}", RowNumber, Cells.Count);
        }

        internal Cell AddCell (string column, string value, CellType type)
        {
            var cell = new Cell () {
                Column = column,
                Row = RowNumber,
                Value = value,
                Type = type
            };

            Cells.Add (column, cell);

            return cell;
        }
    }
}
