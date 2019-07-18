using ExcelTrans.Commands;
using ExcelTrans.Services;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelTrans
{
    public class IntegrationTest
    {
        static Tuple<Stream, string, string> MakeInvoiceFile(IEnumerable<TestData> myData)
        {
            var transform = ExcelService.Encode(new List<IExcelCommand>
            {
                new WorksheetsAdd("Invoice"),
                new CellsStyle(Address.Range, 0, 1, 2, 1, "lc:Yellow"),
            });

            var s = new MemoryStream();
            var w = new StreamWriter(s);
            // add transform to output
            w.WriteLine(transform);
            // add csv file to output
            CsvWriter.Write(w, myData);
            w.Flush(); s.Position = 0;
            var result = new Tuple<Stream, string, string>(s, "text/csv", "invoice.csv");
            // optionally transform
            result = ExcelService.Transform(result);
            return result;
        }

        static void TransferFile(string path, Stream stream, string file)
        {
            path = Path.Combine(path, file);
            if (!Directory.Exists(Path.GetDirectoryName(path)))
                Directory.CreateDirectory(Path.GetDirectoryName(path));
            using (var fileStream = File.Create(path))
            {
                stream.CopyTo(fileStream);
                stream.Seek(0, SeekOrigin.Begin);
            }
        }

        class TestData
        {
            public string One { get; set; }
            public string Two { get; set; }
        }

        [Test]
        public void Should_run()
        {
            var path = @"Out";
            var myData = new[] {
                new TestData { One = "value1", Two = "value2" },
                new TestData { One = "value1", Two = "value2" },
            };
            var file = MakeInvoiceFile(myData);
            TransferFile(path, file.Item1, file.Item3);
        }
    }
}
