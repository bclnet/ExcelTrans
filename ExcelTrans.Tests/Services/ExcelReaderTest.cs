using NFluent;
using NUnit.Framework;
using System.Linq;

namespace ExcelTrans.Services
{
    public class ExcelReaderTest
    {
        string _rawFilePath;
        string _openFilePath;
        string _binaryFilePath;

        [SetUp]
        public void Configure()
        {
            _rawFilePath = "Files/RawTest.xls";
            _openFilePath = "Files/OpenTest.xlsx";
            _binaryFilePath = "Files/BinaryTest.xls";
        }

        [Test]
        public void Should_read_raw()
        {
            // given
            var filePath = _rawFilePath;
            // when
            var doc = ExcelReader.ReadRawXml(filePath, x => x.ToArray(), 9).ToList();
            // then
            Check.That(doc).CountIs(9);
            Check.That(doc.ElementAt(0)).IsEquivalentTo(new[] { "Posting Date", "Tran Date", "Account", "Authorization Number", "Employee Last name", "Employee First name", "Supplier", "Issuer Reference", "Amount USD" });
            Check.That(doc.ElementAt(1)).IsEquivalentTo(new[] { "07/01/2019", "06/30/2019", "XXXX-XXXX-XXXX-9053", "030533", "Roland", "Chris", "Flock - Consultant", "74083429181000001830113", "22.98" });
        }

        [Test]
        public void Should_read_raw2()
        {
            // given
            var filePath = _rawFilePath;
            // when
            var doc = ExcelReader.ReadRaw2Xml(filePath, x => x.ToArray(), 9).ToList();
            // then
            Check.That(doc).CountIs(9);
            Check.That(doc.ElementAt(0)).IsEquivalentTo(new[] { "Posting Date", "Tran Date", "Account", "Authorization Number", "Employee Last name", "Employee First name", "Supplier", "Issuer Reference", "Amount USD" });
            Check.That(doc.ElementAt(1)).IsEquivalentTo(new[] { "07/01/2019", "06/30/2019", "XXXX-XXXX-XXXX-9053", "030533", "Roland", "Chris", "Flock - Consultant", "74083429181000001830113", "22.98" });
        }

        [Test]
        public void Should_read_open()
        {
            // given
            var filePath = _openFilePath;
            // when
            var doc = ExcelReader.ReadOpenXml(filePath, x => x.ToArray(), 2).ToList();
            // then
            Check.That(doc).CountIs(2);
            Check.That(doc.ElementAt(0)).IsEquivalentTo(new[] { "One", "Two" });
            Check.That(doc.ElementAt(1)).IsEquivalentTo(new[] { "value1", "value2" });
        }

        [Test]
        public void Should_read_binary()
        {
            // given
            var filePath = _binaryFilePath;
            // when
            var doc = ExcelReader.ReadBinary(filePath, x => x.ToArray(), 2).ToList();
            // then
            Check.That(doc).CountIs(2);
            Check.That(doc.ElementAt(0)).IsEquivalentTo(new[] { "One", "Two" });
            Check.That(doc.ElementAt(1)).IsEquivalentTo(new[] { "value1", "value2" });
        }
    }
}
