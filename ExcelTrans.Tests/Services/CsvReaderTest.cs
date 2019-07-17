using NFluent;
using NUnit.Framework;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelTrans.Services
{
    public class CsvReaderTest
    {
        MemoryStream _standardStream;
        MemoryStream _pipeDelimitedStream;

        [SetUp]
        public void Configure()
        {
            _standardStream = new MemoryStream(Encoding.ASCII.GetBytes(
@"One,Two
value1,value2"));
            _pipeDelimitedStream = new MemoryStream(Encoding.ASCII.GetBytes(
@"One|Two
value1|value2"));
        }

        [Test]
        public void Should_read_normally()
        {
            // given
            var stream = _standardStream;
            // when
            var doc = CsvReader.Read(stream, x => x.ToArray()).ToList();
            // then
            Check.That(doc).CountIs(2);
            Check.That(doc.ElementAt(0)).IsEquivalentTo(new[] { "One", "Two" });
            Check.That(doc.ElementAt(1)).IsEquivalentTo(new[] { "value1", "value2" });
        }

        [Test]
        public void Should_read_pipe_delimited()
        {
            // given
            var stream = _pipeDelimitedStream;
            // when
            var doc = CsvReader.Read(stream, x => x.ToArray(), settings: new CsvReaderSettings { Delimiter = "|" }).ToList();
            // then
            Check.That(doc).CountIs(2);
            Check.That(doc.ElementAt(0)).IsEquivalentTo(new[] { "One", "Two" });
            Check.That(doc.ElementAt(1)).IsEquivalentTo(new[] { "value1", "value2" });
        }

        [Test]
        public void Should_read_skip_header()
        {
            // given
            var stream = _standardStream;
            // when
            var doc = CsvReader.Read(stream, x => x.ToArray(), startRow: 1).ToList();
            // then
            Check.That(doc).CountIs(1);
            Check.That(doc.ElementAt(0)).IsEquivalentTo(new[] { "value1", "value2" });
        }
    }
}
