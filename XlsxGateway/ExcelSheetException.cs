using System;

namespace XlsxGateway
{
    public class ExcelSheetException : Exception
    {
        public ExcelSheetException(string message) : base(message)
        {
        }
    }
}
