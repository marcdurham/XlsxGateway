using System.Text.RegularExpressions;

namespace XlsxGateway.Models
{

    public class CellAddress
    {
        public static CellAddress From (string cellAddress)
        {
            var match = Regex.Match (cellAddress, @"([a-zA-Z]+)(\d+)");

            string column = "";
            string rowString = "";
            int rowNumber = 0;

            if (!match.Success || match.Groups.Count != 3)
                throw new ExcelSheetException ("Invalid cell address: " + cellAddress);

            column = match.Groups [1].Value;
            rowString = match.Groups [2].Value;
            int.TryParse (rowString, out rowNumber);

            return new CellAddress () {
                Column = column,
                Row = rowNumber,
            };
        }

        public string Column;
        public int Row;
    }
}

