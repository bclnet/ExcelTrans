using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace ExcelTrans.Services
{
    /// <summary>
    /// Processes the input XLS
    /// </summary>
    public static class ExcelReader
    {
        /// <summary>
        /// Reads an Excel.xlsx file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">The file path.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="worksheetPosition">The worksheet position.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">filePath</exception>
        public static IEnumerable<T> ReadOpenXml<T>(string filePath, Func<Collection<string>, T> action, int width = -1, int startRow = 0, string worksheetName = null, int worksheetPosition = 0) => ReadOpenXml(null, new FileInfo(filePath ?? throw new ArgumentNullException(nameof(filePath))), action, width, startRow, worksheetName = null, worksheetPosition = 0);
        /// <summary>
        /// Reads an Excel.xlsx file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="worksheetPosition">The worksheet position.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public static IEnumerable<T> ReadOpenXml<T>(Stream stream, Func<Collection<string>, T> action, int width = -1, int startRow = 0, string worksheetName = null, int worksheetPosition = 0) => ReadOpenXml(stream ?? throw new ArgumentNullException(nameof(stream)), null, action, width, startRow, worksheetName = null, worksheetPosition = 0);
        static IEnumerable<T> ReadOpenXml<T>(Stream stream, FileInfo fileInfo, Func<Collection<string>, T> action, int width = -1, int startRow = 0, string worksheetName = null, int worksheetPosition = 0)
        {
            Collection<string> ParseIntoEntries(Collection<string> list, ExcelRange row)
            {
                list.Clear();
                foreach (var r in row)
                    list.Add(r.Value?.ToString().Trim());
                return list;
            }
            using (var p = fileInfo != null ? new ExcelPackage(fileInfo) : new ExcelPackage(stream ?? throw new ArgumentNullException(nameof(stream))))
            {
                var ws = worksheetName != null ? p.Workbook.Worksheets[worksheetName] : p.Workbook.Worksheets[worksheetPosition];
                var dim = ws.Dimension;
                if (width == -1)
                    width = dim.Columns;
                ExcelRange row = null;
                var list = new Collection<string>();
                for (var rowIdx = startRow + 1; rowIdx <= dim.Rows; rowIdx++)
                    if ((row = ws.Cells[rowIdx, 1, rowIdx, width]) != null)
                        yield return action(ParseIntoEntries(list, row));
            }
        }

        /// <summary>
        /// Reads an Excel.xls file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">The file path.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="worksheetPosition">The worksheet position.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">filePath</exception>
        public static IEnumerable<T> ReadBinary<T>(string filePath, Func<Collection<string>, T> action, int width, int startRow = 0, string worksheetName = null, int worksheetPosition = 0) => ReadBinary(File.OpenRead(filePath ?? throw new ArgumentNullException(nameof(filePath))), action, width, startRow, worksheetName, worksheetPosition);
        /// <summary>
        /// Reads an Excel.xls file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <param name="worksheetName">Name of the worksheet.</param>
        /// <param name="worksheetPosition">The worksheet position.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public static IEnumerable<T> ReadBinary<T>(Stream stream, Func<Collection<string>, T> action, int width, int startRow = 0, string worksheetName = null, int worksheetPosition = 0)
        {
            Collection<string> ParseIntoEntries(Collection<string> list, IRow r)
            {
                list.Clear();
                for (var i = 0; i < width; i++)
                    list.Add(r.GetCell(i).StringCellValue?.Trim());
                return list;
            }
            try
            {
                var p = new HSSFWorkbook(stream ?? throw new ArgumentNullException(nameof(stream)));
                {
                    var ws = worksheetName != null ? p.GetSheet(worksheetName) : p.GetSheetAt(worksheetPosition);
                    IRow row = null;
                    var list = new Collection<string>();
                    for (var rowIdx = startRow; rowIdx <= ws.LastRowNum; rowIdx++)
                        if ((row = ws.GetRow(rowIdx)) != null)
                            yield return action(ParseIntoEntries(list, row));
                }
            }
            finally { stream.Dispose(); }
        }

        /// <summary>
        /// Reads an raw xml Excel file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">The file path.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">filePath</exception>
        public static IEnumerable<T> ReadRawXml<T>(string filePath, Func<Collection<string>, T> action, int width, int startRow = 0) => ReadRawXml(File.OpenRead(filePath ?? throw new ArgumentNullException(nameof(filePath))), action, width, startRow);
        /// <summary>
        /// Reads an raw xml Excel file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public static IEnumerable<T> ReadRawXml<T>(Stream stream, Func<Collection<string>, T> action, int width, int startRow = 0)
        {
            Exception ParsingException(Exception e, string xml)
            {
                var msg = e.Message;
                if (!msg.Contains("Line") || !msg.Contains("position"))
                    return e;
                var pidx = msg.IndexOf("Line"); var pidx2 = msg.IndexOf(",", pidx); var line = int.Parse(msg.Substring(pidx + 4, pidx2 - pidx - 4));
                if (line != 1)
                    return e;
                pidx = msg.IndexOf("position"); pidx2 = msg.IndexOf(".", pidx); var position = int.Parse(msg.Substring(pidx + 8, pidx2 - pidx - 8));
                var error = xml.Substring(position - 30, 30) + "!!" + xml.Substring(position, 20);
                return new ArgumentOutOfRangeException(msg, error);
            }
            Collection<string> ParseIntoEntries(Collection<string> list, XElement row)
            {
                var cols = row.Descendants("th").Concat(row.Descendants("td")).ToArray();
                list.Clear();
                for (var i = 0; i < width; i++)
                    list.Add(i < cols.Length ? cols[i].Value?.Trim() : null);
                return list;
            }
            var xml_ = new StreamReader(stream ?? throw new ArgumentNullException(nameof(stream))).ReadToEnd();
            int idx = 0, idx2;
            while (true)
            {
                idx = xml_.IndexOf("<table", idx);
                idx2 = idx != -1 ? xml_.IndexOf("</table>", idx) : -1;
                if (idx2 == -1)
                    break;
                var xml = xml_.Substring(idx, idx2 - idx + 8).Replace(" nowrap", "").Replace("&", "&amp;");
                XDocument doc;
                try
                {
                    using (var s = new StringReader(xml))
                        doc = XDocument.Load(s);
                }
                catch (Exception e) { throw ParsingException(e, xml); }
                var list = new Collection<string>();
                foreach (var row in doc.Descendants("tr"))
                {
                    if (startRow > 0)
                    {
                        startRow--;
                        continue;
                    }
                    var entries = ParseIntoEntries(list, row);
                    yield return action(entries);
                }
                // next
                idx = idx2;
            }
        }

        /// <summary>
        /// Reads an raw xml Excel file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">The file path.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">filePath</exception>
        public static IEnumerable<T> ReadRaw2Xml<T>(string filePath, Func<Collection<string>, T> action, int width, int startRow = 0) => ReadRaw2Xml(File.OpenRead(filePath ?? throw new ArgumentNullException(nameof(filePath))), action, width, startRow);
        /// <summary>
        /// Reads an raw xml Excel file.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">The stream.</param>
        /// <param name="action">The action.</param>
        /// <param name="width">The width.</param>
        /// <param name="startRow">The start row.</param>
        /// <returns>IEnumerable&lt;T&gt;.</returns>
        /// <exception cref="System.ArgumentNullException">stream</exception>
        public static IEnumerable<T> ReadRaw2Xml<T>(Stream stream, Func<Collection<string>, T> action, int width, int startRow = 0)
        {
            Collection<string> ParseIntoEntries(Collection<string> list, XmlReader row)
            {
                list.Clear();
                while (row.Read())
                {
                    if (row.NodeType == XmlNodeType.EndElement && row.Name == "tr")
                    {
                        for (var i = list.Count; i < width; i++)
                            list.Add(null);
                        return list;
                    }
                    if (row.NodeType != XmlNodeType.Text)
                        continue;
                    list.Add(row.Value.Trim());
                }
                throw new InvalidOperationException();
            }
            var xml_ = new StreamReader(stream ?? throw new ArgumentNullException(nameof(stream))).ReadToEnd();
            int idx = 0, idx2;
            while (true)
            {
                idx = xml_.IndexOf("<table", idx);
                idx2 = idx != -1 ? xml_.IndexOf("</table>", idx) : -1;
                if (idx2 == -1)
                    break;
                var xml = xml_.Substring(idx, idx2 - idx + 8).Replace(" nowrap", "").Replace("&", "&amp;");
                var list = new Collection<string>();
                using (var row = XmlReader.Create(new StringReader(xml)))
                {
                    row.MoveToContent();
                    while (row.Read())
                    {
                        if (row.NodeType != XmlNodeType.Element || row.Name != "tr")
                            continue;
                        if (startRow > 0)
                        {
                            startRow--;
                            continue;
                        }
                        var entries = ParseIntoEntries(list, row);
                        yield return action(entries);
                    }
                }
                // next
                idx = idx2;
            }
        }
    }
}