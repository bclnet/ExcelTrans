using System.Linq;
using ExcelTrans.Commands;
using NFluent;
using NUnit.Framework;

namespace ExcelTrans
{
    public class ExcelServiceTest
    {
        IExcelCommand[] _simpleCmds;
        IExcelContext _excelContext;

        [SetUp]
        public void Configure()
        {
            _simpleCmds = new IExcelCommand[] { new CellsStyle("A1", "f1") };
            _excelContext = new ExcelContext();
        }

        [TearDown]
        public void TearDown()
        {
            _excelContext.Dispose();
        }

        [Test]
        public void Should_encode_with_description()
        {
            // given
            var cmds = _simpleCmds;
            // when
            var doc = ExcelService.Encode(true, cmds);
            // then
            Check.That(doc).IsEqualTo(
@"^q|   CellsStyle[A1]: f1
^q=AQAAAAJBMQEAAmYx");
        }

        [Test]
        public void Should_encode_without_description()
        {
            // given
            var cmds = _simpleCmds;
            // when
            var doc = ExcelService.Encode(false, cmds);
            // then
            Check.That(doc).IsEqualTo("^q=AQAAAAJBMQEAAmYx");
        }

        [Test]
        public void Should_decode_correctly()
        {
            // given
            var value = "^q=AQAAAAJBMQEAAmYx";
            var value2 = ExcelService.Encode(false, _simpleCmds);
            // when
            var cmds = ExcelService.Decode(value);
            var cmds2 = ExcelService.Decode(value2);
            // then
            Check.That(cmds.Select(x => ((CellsStyle)x).Cells)).IsEquivalentTo(_simpleCmds.Select(x => ((CellsStyle)x).Cells));
            Check.That(cmds2.Select(x => ((CellsStyle)x).Cells)).IsEquivalentTo(_simpleCmds.Select(x => ((CellsStyle)x).Cells));
        }

        [Test]
        public void Should_getaddress_correctly()
        {
            Check.That(ExcelService.GetAddressCol(1)).IsEqualTo("A");
            Check.That(ExcelService.GetAddressRow(2)).IsEqualTo("2");
            Check.That(ExcelService.GetAddress(03, "A")).IsEqualTo("A3");
            Check.That(ExcelService.GetAddress(04, 1)).IsEqualTo("A4");
            Check.That(ExcelService.GetAddress(05, true, "A", true)).IsEqualTo("$A$5"); Check.That(ExcelService.GetAddress(05, false, "A", false)).IsEqualTo("A5");
            Check.That(ExcelService.GetAddress(06, true, 1, true)).IsEqualTo("$A$6"); Check.That(ExcelService.GetAddress(06, false, 1, false)).IsEqualTo("A6");
            Check.That(ExcelService.GetAddress(07, "A", true)).IsEqualTo("$A$7"); Check.That(ExcelService.GetAddress(07, "A", false)).IsEqualTo("A7");
            Check.That(ExcelService.GetAddress(08, 1, true)).IsEqualTo("$A$8"); Check.That(ExcelService.GetAddress(08, 1, false)).IsEqualTo("A8");
            Check.That(ExcelService.GetAddress(09, "A", 09, "A")).IsEqualTo("A9"); Check.That(ExcelService.GetAddress(09, "A", 09, "B")).IsEqualTo("A9:B9");
            Check.That(ExcelService.GetAddress(10, 1, 10, 1)).IsEqualTo("A10"); Check.That(ExcelService.GetAddress(10, 1, 10, 2)).IsEqualTo("A10:B10");
            Check.That(ExcelService.GetAddress(11, "A", 11, "B", true)).IsEqualTo("$A$11:$B$11");
            Check.That(ExcelService.GetAddress(12, 1, 12, 2, true)).IsEqualTo("$A$12:$B$12");
            Check.That(ExcelService.GetAddress(13, "A", 13, "B", true, true, true, true)).IsEqualTo("$A$13:$B$13");
            Check.That(ExcelService.GetAddress(14, 1, 14, 2, true, true, true, true)).IsEqualTo("$A$14:$B$14");
            Check.That(ExcelService.GetAddress(Address.Cell, 15, "A")).IsEqualTo("^17:15:1");
            Check.That(ExcelService.GetAddress(Address.Cell, 16, 1)).IsEqualTo("^17:16:1");
            Check.That(ExcelService.GetAddress(Address.Range, 17, "A", 17, "B")).IsEqualTo("^18:17:1:17:2");
            Check.That(ExcelService.GetAddress(Address.Range, 18, 1, 18, 2)).IsEqualTo("^18:18:1:18:2");
            Check.That(ExcelService.GetAddress(_excelContext, Address.Cell, 19, "A")).IsEqualTo("B20");
            Check.That(ExcelService.GetAddress(_excelContext, Address.Cell, 20, 1)).IsEqualTo("B21");
            Check.That(ExcelService.GetAddress(_excelContext, Address.Range, 21, "A", 21, "B")).IsEqualTo("B22:C22");
            Check.That(ExcelService.GetAddress(_excelContext, Address.Range, 22, 1, 22, 2)).IsEqualTo("B23:C23");
        }

        [Test]
        public void Should_decodeaddress_correctly()
        {
            Check.That(ExcelService.DecodeAddress(_excelContext, "^17:15:1")).IsEqualTo("B16");
            Check.That(ExcelService.DecodeAddress(_excelContext, "^18:17:1:17:2")).IsEqualTo("B18:C18");
        }

        [Test]
        public void Should_transform_correctly()
        {
        }
    }
}
