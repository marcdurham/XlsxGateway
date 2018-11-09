using System.IO;

namespace XlsxGateway.Gateways
{
    public interface ISharedStringGateway
    {
        void OpenFrom (string streamText);
        void Add (string value);
        void Save (Stream targetSharedStringsStream);
        string StringAtIndexOf (string value);
        int Count { get; }
    }
}

