using NUnit.Framework;
using XlsxGateway.Gateways;
using System.Xml;
using XlsxGateway.Models;

namespace XlsxGateway.UnitTests
{
    [TestFixture]
    public class SheetDocumentXmlReaderTests
    {
        [Test]
        public void CellFrom_NullInput_ReturnsEmptyCell()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            Assert.That(gateway.CellFrom(null).IsEmpty, Is.True);
        }

        [Test]
        public void CellFrom_CellAddress_A1_ReturnsColumn()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = CellElement(address: "A1", value: "1");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Column, Is.EqualTo("A"));
        }

        [Test]
        public void CellFrom_CellAddress_A1_ReturnsRow()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = CellElement(address: "A1", value: "1");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Row, Is.EqualTo(1));
        }

        [Test]
        public void CellFrom_MissingTypeAttribute_ReturnsNumberType()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = CellElement(address: "A1", value: "1");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Type, Is.EqualTo(CellType.Number));
        }

        [Test]
        public void CellFrom_NumberTypeExplicit_ReturnsNumberType()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = CellElement(address: "A1", value: "2");
            cellNode.SetAttribute("t", "n");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Type, Is.EqualTo(CellType.Number));
        }

        [Test]
        public void CellFrom_StyleMissing_ReturnsStyleNull()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = CellElement(address: "A1", value: "2");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Style, Is.Null);
        }

        [Test]
        public void CellFrom_Style2_ReturnsStyle2()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = CellElement(address: "A1", value: "2");
            cellNode.SetAttribute("s", "2");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Style, Is.EqualTo(2));
        }

        [Test]
        public void CellFrom_InlineString_HasStringValue()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = InlineStringCellElement(address: "A1", value: "test string");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Value, Is.EqualTo("test string"));
        }

        [Test]
        public void CellFrom_InlineString_HasInlineStringType()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = InlineStringCellElement(address: "A1", value: "test string");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Type, Is.EqualTo(CellType.InlineString));
        }

        [Test]
        public void CellFrom_SharedString_HasValue()
        {
            var sharedStringGateway = new SharedStringGatewayFake();
            sharedStringGateway.CountReturns = 123;
            sharedStringGateway.StringAtIndexOfReturns = "real test string";
             
            var gateway = new SheetDocumentXmlReader(sharedStringGateway);

            XmlElement cellNode = SharedStringCellElement(address: "A1", value: "123");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Value, Is.EqualTo("real test string"));
            Assert.That(sharedStringGateway.StringAtIndexOfValueParameter, Is.EqualTo("123"));
        }


        [Test]
        public void CellFrom_SharedString_HasSharedStringType()
        {
            SheetDocumentXmlReader gateway = DefaultGateway();

            XmlElement cellNode = SharedStringCellElement(address: "A1", value: "test string");

            Cell cell = gateway.CellFrom(cellNode);

            Assert.That(cell.Type, Is.EqualTo(CellType.SharedString));
        }

        private static XmlElement CellElement(string address, string value)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement cellNode = doc.CreateElement("c");
            XmlNode valueNode = doc.CreateElement("v");
            valueNode.InnerText = value;
            cellNode.AppendChild(valueNode);
            cellNode.SetAttribute("r", address);
            return cellNode;
        }

        private static XmlElement InlineStringCellElement(string address, string value)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement cellNode = doc.CreateElement("c");
            XmlNode innerNode = doc.CreateElement("is");
            XmlNode valueNode = doc.CreateElement("t");
            valueNode.InnerText = value;
            innerNode.AppendChild(valueNode);
            cellNode.AppendChild(innerNode);
            cellNode.SetAttribute("r", address);
            cellNode.SetAttribute("t", "inlineStr");
            return cellNode;
        }

        private static XmlElement SharedStringCellElement(string address, string value)
        {
            XmlDocument doc = new XmlDocument();
            XmlElement cellNode = doc.CreateElement("c");
            XmlNode valueNode = doc.CreateElement("v");
            valueNode.InnerText = value;
            cellNode.AppendChild(valueNode);
            cellNode.SetAttribute("r", address);
            cellNode.SetAttribute("t", "s");
            return cellNode;
        }

        private static SheetDocumentXmlReader DefaultGateway()
        {
            var sharedStringGateway = new SharedStringGatewayFake();
            var gateway = new SheetDocumentXmlReader(sharedStringGateway);
            return gateway;
        }
    }
}
