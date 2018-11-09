using System.Linq;
using XlsxGateway.Models;

namespace JXlsxGateway.Tools
{
    public class RowCopier
    {
        private Worksheet sourceSheet;
        private Worksheet targetSheet;
        private int targetRowIndex;

        public RowCopier(Worksheet sourceSheet, Worksheet targetSheet)
        {
            this.sourceSheet = sourceSheet;
            this.targetSheet = targetSheet;
        }

        /// <summary>
        /// Copies an entire row from the source spreadsheet to the target.
        /// </summary>
        /// <param name="sourceRowIndex">Source row index starts at zero, and includes the header.</param>
        /// <param name="targetRowIndex">Target row index starts at zero, and includes the header.</param>
        public void Copy(int sourceRowIndex, int targetRowIndex)
        {
            this.targetRowIndex = targetRowIndex;

            // Row is copied, but rowNumber is changed
            Row sourceRow = sourceSheet.Row(sourceRowIndex);
	
            CopyCells (from: sourceRow, toRowNumber: RowNumberFrom(targetRowIndex));
        }
        
        private void CopyCells(Row from, int toRowNumber)
        {
            var newRow = targetSheet.AddRow (toRowNumber);
            var cells = from.Cells.Select (c => c.Value);
            
            foreach (Cell cell in cells) 
                newRow.AddCell (column: cell.Column, value: cell.Value, type: cell.Type);

        }

        private int RowNumberFrom (int rowIndex)
        {
            return rowIndex + 1;
        }
    }
}
