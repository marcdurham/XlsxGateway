namespace XlsxGateway.Tools
{
    public class ColumnNamer
    {
        public static string NameOf(int index)
        {
            string range = "";
            if (index < 0)
                return range;

            for (int i = 1; index + i > 0; i = 0)
            {
                range = ((char)(65 + index % 26)) + range;
                index /= 26;
            }

            if (range.Length > 1)
                range = ((char)(range[0] - 1)) + range.Substring(1);

            return range;
        }
    }
}
