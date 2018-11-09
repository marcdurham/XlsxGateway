using System.IO;

namespace XlsxGateway.SystemAdapters
{
    public class CompressedEntry
    {
        public string Name;
        public string Content;
        public Stream Open()
        {
            return null;
        }
    }
}
