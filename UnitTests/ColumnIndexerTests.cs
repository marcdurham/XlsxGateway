using NUnit.Framework;
using XlsxGateway.Tools;

namespace XlsxGateway.UnitTests
{
    public class ColumnIndexerTests
    {
        [Test]
        public void NameOf_1_A()
        {
            Assert.AreEqual("A", ColumnNamer.NameOf(0));
        }

        [Test]
        public void NameOf_26_AA()
        {
            Assert.AreEqual("AA", ColumnNamer.NameOf(26));
        }

        [Test]
        public void NameOf_27_AB()
        {
            Assert.AreEqual("AB", ColumnNamer.NameOf(27));
        }

        [Test]
        public void NameOf_676_AAA()
        {
            Assert.AreEqual("AAA", ColumnNamer.NameOf(676));
        }
    }
}
