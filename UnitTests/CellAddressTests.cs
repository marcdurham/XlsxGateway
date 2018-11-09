using NUnit.Framework;
using XlsxGateway.Models;

namespace XlsxGateway.UnitTests
{
    [TestFixture]
    public class CellAddressTests
    {
        [Test]
        public void From_B3_Returns_ColumnB ()
        {
            Assert.That (CellAddress.From ("B3").Column, Is.EqualTo ("B"));
        }

        [Test]
        public void From_B3_Returns_Row3 ()
        {
            Assert.That (CellAddress.From ("B3").Row, Is.EqualTo (3));
        }

        [Test]
        public void From_C23_Returns_Row23 ()
        {
            Assert.That (CellAddress.From ("C23").Row, Is.EqualTo (23));
        }

        [Test]
        public void From_C23_Returns_ColumnC ()
        {
            Assert.That (CellAddress.From ("C23").Column, Is.EqualTo ("C"));
        }

        [Test]
        public void From_ABC987_Returns_ColumnABC ()
        {
            Assert.That (CellAddress.From ("ABC987").Column, Is.EqualTo ("ABC"));
        }

        [Test]
        public void From_ABC987_Returns_Row987 ()
        {
            Assert.That (CellAddress.From ("ABC987").Row, Is.EqualTo (987));
        }
    }
}

