using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using WriteOptions = ExcelTrans.Services.CsvWriterOptions.WriteOptions;

namespace ExcelTrans.Services
{
    /// <summary>
    /// CsvWriter
    /// </summary>
    public static class CsvWriter
    {
        /// <summary>
        /// Writes the specified context.
        /// </summary>
        /// <typeparam name="TItem">The type of the item.</typeparam>
        /// <param name="w">The w.</param>
        /// <param name="set">The set.</param>
        /// <param name="options">The context.</param>
        public static void Write<TItem>(TextWriter w, IEnumerable<TItem> set, CsvWriterOptions options = null)
        {
            if (w == null)
                throw new ArgumentNullException(nameof(w));
            if (set == null)
                throw new ArgumentNullException(nameof(set));

            if (options == null)
                options = CsvWriterOptions.Default;
            var delimiter = options.Delimiter[0];
            var hasHeaderRow = (options.EmitOptions & WriteOptions.HasHeaderRow) != 0;
            var encodeValues = (options.EmitOptions & WriteOptions.EncodeValues) != 0;
            var columns = options.GetColumns != null ? options.GetColumns(typeof(TItem)) : CsvWriterOptions.GetColumnsByType(typeof(TItem), hasHeaderRow);

            // header
            var fields = options.Fields.Count > 0 ? options.Fields : null;
            var b = new StringBuilder();
            if (hasHeaderRow)
            {
                foreach (var column in columns)
                {
                    // label
                    var name = column.DisplayName ?? column.Name;
                    if (fields != null && fields.TryGetValue(column.Name, out var field) && field != null)
                        if (field.Ignore) continue;
                        else if (field.DisplayName != null) name = field.DisplayName;
                    b.Append(Encode(encodeValues ? EncodeValue(name) : name) + delimiter);
                }
                if (b.Length > 0)
                    b.Length--;
                w.Write(b.ToString() + Environment.NewLine);
            }

            // rows
            try
            {
                foreach (var group in set.Cast<object>().GroupAt(options.FlushAt))
                {
                    var newGroup = options.BeforeFlush == null ? group : options.BeforeFlush(group);
                    if (newGroup == null)
                        return;
                    foreach (var item in newGroup)
                    {
                        b.Length = 0;
                        foreach (var column in columns)
                        {
                            // value
                            string value;
                            var itemValue = column.GetValue(item);
                            if (fields != null && fields.TryGetValue(column.Name, out var field) && field != null)
                            {
                                if (field.Ignore) continue;
                                value = field.CustomFieldFormatter == null ? itemValue?.ToString() ?? string.Empty : field.CustomFieldFormatter(field, item, itemValue);
                                if (value.Length == 0)
                                    value = field.DefaultValue ?? string.Empty;
                                if (value.Length != 0)
                                {
                                    var args = field.Args;
                                    if (args != null)
                                    {
                                        if (args.doNotEncode == true)
                                        {
                                            b.Append(value + delimiter);
                                            continue;
                                        }
                                        if (args.asExcelFunction == true)
                                            value = "=" + value;
                                    }
                                }
                            }
                            else value = itemValue?.ToString() ?? string.Empty;
                            // append value
                            b.Append(Encode(encodeValues ? EncodeValue(value) : value) + delimiter);
                        }
                        if (b.Length > 0)
                            b.Length--;
                        w.Write(b.ToString() + Environment.NewLine);
                    }
                    w.Flush();
                }
            }
            finally { w.Flush(); }
        }

        static string Encode(string value) => value;
        static string EncodeValue(string value) => string.IsNullOrEmpty(value) ? "\"\"" : "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}