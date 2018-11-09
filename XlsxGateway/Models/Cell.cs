namespace XlsxGateway.Models
{
    public class Cell
    {
        public static Cell Empty = new EmptyCell();

        public string Column;
        public int Row;
        public string Value;
        public CellType Type;
        public int? Style;

        public virtual bool IsEmpty { get { return false; } }

        public string Address
        {
            get { return (Column == null ? null : Column + Row); }
        }

        public override string ToString()
        {
            return string.Format("{0}{1}:{2}", Column, Row, Value);
        }
    }

    public class EmptyCell : Cell
    {
        public override bool IsEmpty { get { return true; } }

        public override string ToString()
        {
            return string.Format("{0}{1}:<Emtpy>", Column, Row);
        }
    }

    public enum CellType
    {
        SharedString,
        Number,
        InlineString
    }

}
