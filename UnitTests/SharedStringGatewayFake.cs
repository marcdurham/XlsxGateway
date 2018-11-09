using System.IO;
using XlsxGateway.Gateways;

namespace XlsxGateway.UnitTests
{
    public class SharedStringGatewayFake : ISharedStringGateway
    {
        public int CountReturns;
        public int Count { get { return CountReturns; } }

        public string AddValueParameter;
        public void Add(string value)
        {
            AddValueParameter = value;
        }

        public void OpenFrom(string streamText)
        {
        }

        public void Save(Stream targetSharedStringsStream)
        {
        }

        public string StringAtIndexOfReturns;
        public string StringAtIndexOfValueParameter;
        public string StringAtIndexOf(string value)
        {
            StringAtIndexOfValueParameter = value;
            return StringAtIndexOfReturns;
        }
    }
}