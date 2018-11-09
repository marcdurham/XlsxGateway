using NUnit.Framework;
using XlsxGateway.Models;

namespace XlsxGateway.UnitTests
{
    [TestFixture]
    public class CellTests
    {
        [Test]
        public void Address_ColumnNull_RowNull_ReturnsNull ()
        {
            Cell cell = new Cell ();
            Assert.That (cell.Address, Is.Null);
        }

        [Test]
        public void Address_ColumnA_Row1_ReturnsA1 ()
        {
            Cell cell = new Cell { Column = "A", Row = 1 };
            Assert.That (cell.Address, Is.EqualTo ("A1"));
        }
    }
}

