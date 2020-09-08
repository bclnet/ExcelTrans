using NFluent;
using NUnit.Framework;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelTrans.Services
{
    public class CsvWriterTest
    {
        IEnumerable<TestItem> _testItems;

        class TestItem
        {
            public string One { get; set; }
            public string Two { get; set; }
        }

        [SetUp]
        public void Configure()
        {
            _testItems = new[] { new TestItem { One = "value1", Two = "value2" } };
        }

        [Test]
        public void Should_write_normally()
        {
            // given
            var b = new StringBuilder();
            // when
            CsvWriter.Write(new StringWriter(b), _testItems);
            var doc = b.ToString();
            // then
            Check.That(doc).IsEqualTo(
@"""One"",""Two""
""value1"",""value2""
".ToLocalString());
        }

        [Test]
        public void Should_write_pipe_delimited()
        {
            // given
            var b = new StringBuilder();
            // when
            CsvWriter.Write(new StringWriter(b), _testItems, options: new CsvWriterOptions { Delimiter = "|" });
            var doc = b.ToString();
            // then
            Check.That(doc).IsEqualTo(
@"""One""|""Two""
""value1""|""value2""
".ToLocalString());
        }


        [Test]
        public void Should_write_skip_header()
        {
            // given
            var b = new StringBuilder();
            // when
            CsvWriter.Write(new StringWriter(b), _testItems, options: new CsvWriterOptions { HasHeaderRow = false });
            var doc = b.ToString();
            // then
            Check.That(doc).IsEqualTo(
@"""value1"",""value2""
".ToLocalString());
        }
    }
}
