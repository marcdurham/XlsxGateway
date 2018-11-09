namespace XlsxGateway.SystemAdapters
{
    public interface ICompressor
    {
        void OpenForUpdate(string fileName);
        void OpenToRead(string fileName);
        void Save();
        CompressedEntry GetEntry(string name);
    }
}
