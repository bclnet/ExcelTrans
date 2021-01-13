using ExcelTrans.Commands;
using ExcelTrans.Services;
using ExcelTrans.Utils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ExcelTrans
{
    /// <summary>
    /// ExcelService
    /// </summary>
    public static class ExcelService
    {
        public static string Author = "S";
        /// <summary>
        /// The comment
        /// </summary>
        public static readonly string Comment = "^q|";
        /// <summary>
        /// The stream
        /// </summary>
        public static readonly string Stream = "^q=";
        /// <summary>
        /// The break
        /// </summary>
        public static readonly string Break = "^q!";
        /// <summary>
        /// Gets the pop frame.
        /// </summary>
        /// <value>
        /// The pop frame.
        /// </value>
        public static string PopFrame => $"{Stream}{ExcelSerDes.Encode(new PopFrame())}";
        /// <summary>
        /// Gets the pop set.
        /// </summary>
        /// <value>
        /// The pop set.
        /// </value>
        public static string PopSet => $"{Stream}{ExcelSerDes.Encode(new PopSet())}";
        /// <summary>
        /// Encodes the specified describe.
        /// </summary>
        /// <param name="describe">if set to <c>true</c> [describe].</param>
        /// <param name="cmds">The CMDS.</param>
        /// <returns></returns>
        public static string Encode(bool describe, params IExcelCommand[] cmds) => $"{(describe ? ExcelSerDes.Describe(Comment, cmds) : null)}{Stream}{ExcelSerDes.Encode(cmds)}";
        /// <summary>
        /// Encodes the specified describe.
        /// </summary>
        /// <param name="describe">if set to <c>true</c> [describe].</param>
        /// <param name="cmds">The CMDS.</param>
        /// <returns></returns>
        public static string Encode(bool describe, IEnumerable<IExcelCommand> cmds) => Encode(describe, cmds.ToArray());
        /// <summary>
        /// Encodes the specified CMDS.
        /// </summary>
        /// <param name="cmds">The CMDS.</param>
        /// <returns></returns>
        public static string Encode(params IExcelCommand[] cmds) => $"{Stream}{ExcelSerDes.Encode(cmds)}";
        /// <summary>
        /// Encodes the specified CMDS.
        /// </summary>
        /// <param name="cmds">The CMDS.</param>
        /// <returns></returns>
        public static string Encode(IEnumerable<IExcelCommand> cmds) => Encode(cmds.ToArray());
        /// <summary>
        /// Decodes the specified packed.
        /// </summary>
        /// <param name="packed">The packed.</param>
        /// <returns></returns>
        public static IExcelCommand[] Decode(string packed) => ExcelSerDes.Decode(packed.Substring(Stream.Length));

        /// <summary>
        /// Transforms the specified value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="allowMacros">if set to <c>true</c> [allow macros].</param>
        /// <returns></returns>
        public static (Stream stream, string, string path) Transform((Stream stream, string, string path) value, bool macrosDisabled = false)
        {
            using (var s1 = value.stream)
            {
                s1.Seek(0, SeekOrigin.Begin);
                var s2 = new MemoryStream(Build(s1, macrosDisabled, out var macros));
                return (s2, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", value.path.Replace(".csv", !macros ? ".xlsx" : ".xlsm"));
            }
        }

        static byte[] Build(Stream s, bool macrosDisabled, out bool macros)
        {
            using (var ctx = new ExcelContext(macrosDisabled))
            {
                CsvReader.Read(s, x =>
                {
                    if (x == null || x.Count == 0 || x[0].StartsWith(Comment)) return true;
                    else if (x[0].StartsWith(Stream))
                    {
                        var r = ctx.ExecuteCmd(Decode(x[0]), out var action) != null;
                        action?.Invoke();
                        return r;
                    }
                    else if (x[0].StartsWith(Break)) return false;
                    ctx.AdvanceRow();
                    if (ctx.Sets.Count == 0) ctx.WriteRow(x);
                    else ctx.Sets.Peek().Add(x);
                    return true;
                }).Any(x => !x);
                ctx.Flush();
                macros = ctx.V != null;
                return ctx.P.GetAsByteArray();
            }
        }

        /// <summary>
        /// Gets the address col.
        /// </summary>
        /// <param name="column">The column.</param>
        /// <returns></returns>
        public static string GetAddressCol(int column) => ExcelCellBase.GetAddressCol(column);
        /// <summary>
        /// Gets the address row.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        public static string GetAddressRow(int row) => ExcelCellBase.GetAddressRow(row);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <returns></returns>
        public static string GetAddress(int row, string column) => ExcelCellBase.GetAddress(row, ColToInt(column));
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <returns></returns>
        public static string GetAddress(int row, int column) => ExcelCellBase.GetAddress(row, column);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns></returns>
        public static string GetAddress((int row, int column) cell) => ExcelCellBase.GetAddress(cell.row, cell.column);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="absoluteRow">if set to <c>true</c> [absolute row].</param>
        /// <param name="column">The column.</param>
        /// <param name="absoluteCol">if set to <c>true</c> [absolute col].</param>
        /// <returns></returns>
        public static string GetAddress(int row, bool absoluteRow, string column, bool absoluteCol) => ExcelCellBase.GetAddress(row, absoluteRow, ColToInt(column), absoluteCol);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="absoluteRow">if set to <c>true</c> [absolute row].</param>
        /// <param name="column">The column.</param>
        /// <param name="absoluteCol">if set to <c>true</c> [absolute col].</param>
        /// <returns></returns>
        public static string GetAddress(int row, bool absoluteRow, int column, bool absoluteCol) => ExcelCellBase.GetAddress(row, absoluteRow, column, absoluteCol);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="absolute">if set to <c>true</c> [absolute].</param>
        /// <returns></returns>
        public static string GetAddress(int row, string column, bool absolute) => ExcelCellBase.GetAddress(row, ColToInt(column), absolute);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="column">The column.</param>
        /// <param name="absolute">if set to <c>true</c> [absolute].</param>
        /// <returns></returns>
        public static string GetAddress(int row, int column, bool absolute) => ExcelCellBase.GetAddress(row, column, absolute);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <returns></returns>
        public static string GetAddress(int fromRow, string fromColumn, int toRow, string toColumn) => ExcelCellBase.GetAddress(fromRow, ColToInt(fromColumn), toRow, ColToInt(toColumn));
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <returns></returns>
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns></returns>
        public static string GetAddress((int fromRow, int fromColumn, int toRow, int toColumn) cell) => ExcelCellBase.GetAddress(cell.fromRow, cell.fromColumn, cell.toRow, cell.toColumn);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <param name="absolute">if set to <c>true</c> [absolute].</param>
        /// <returns></returns>
        public static string GetAddress(int fromRow, string fromColumn, int toRow, string toColumn, bool absolute) => ExcelCellBase.GetAddress(fromRow, ColToInt(fromColumn), toRow, ColToInt(toColumn), absolute);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <param name="absolute">if set to <c>true</c> [absolute].</param>
        /// <returns></returns>
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn, bool absolute) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, absolute);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <param name="fixedFromRow">if set to <c>true</c> [fixed from row].</param>
        /// <param name="fixedFromColumn">if set to <c>true</c> [fixed from column].</param>
        /// <param name="fixedToRow">if set to <c>true</c> [fixed to row].</param>
        /// <param name="fixedToColumn">if set to <c>true</c> [fixed to column].</param>
        /// <returns></returns>
        public static string GetAddress(int fromRow, string fromColumn, int toRow, string toColumn, bool fixedFromRow, bool fixedFromColumn, bool fixedToRow, bool fixedToColumn) => ExcelCellBase.GetAddress(fromRow, ColToInt(fromColumn), toRow, ColToInt(toColumn), fixedFromRow, fixedFromColumn, fixedToRow, fixedToColumn);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <param name="fixedFromRow">if set to <c>true</c> [fixed from row].</param>
        /// <param name="fixedFromColumn">if set to <c>true</c> [fixed from column].</param>
        /// <param name="fixedToRow">if set to <c>true</c> [fixed to row].</param>
        /// <param name="fixedToColumn">if set to <c>true</c> [fixed to column].</param>
        /// <returns></returns>
        public static string GetAddress(int fromRow, int fromColumn, int toRow, int toColumn, bool fixedFromRow, bool fixedFromColumn, bool fixedToRow, bool fixedToColumn) => ExcelCellBase.GetAddress(fromRow, fromColumn, toRow, toColumn, fixedFromRow, fixedFromColumn, fixedToRow, fixedToColumn);
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <returns></returns>
        public static string GetAddress(Address r, int row, string col) => $"^{(int)r}:{row}:{ColToInt(col)}";
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <returns></returns>
        public static string GetAddress(Address r, int row, int col) => $"^{(int)r}:{row}:{col}";
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="cell">The cell.</param>
        /// <returns></returns>
        public static string GetAddress((Address r, int row, int col) cell) => $"^{(int)cell.r}:{cell.row}:{cell.col}";
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <returns></returns>
        public static string GetAddress(Address r, int fromRow, string fromColumn, int toRow, string toColumn) => $"^{(int)r}:{fromRow}:{ColToInt(fromColumn)}:{toRow}:{ColToInt(toColumn)}";
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <returns></returns>
        public static string GetAddress(Address r, int fromRow, int fromColumn, int toRow, int toColumn) => $"^{(int)r}:{fromRow}:{fromColumn}:{toRow}:{toColumn}";
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns></returns>
        public static string GetAddress((Address r, int fromRow, int fromColumn, int toRow, int toColumn) cell) => $"^{(int)cell.r}:{cell.fromRow}:{cell.fromColumn}:{cell.toRow}:{cell.toColumn}";
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <returns></returns>
        public static string GetAddress(this IExcelContext ctx, Address r, int row, string col) => DecodeAddress(ctx, GetAddress(r, row, col));
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <returns></returns>
        public static string GetAddress(this IExcelContext ctx, Address r, int row, int col) => DecodeAddress(ctx, GetAddress(r, row, col));
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="cell">The cell.</param>
        /// <returns></returns>
        public static string GetAddress(this IExcelContext ctx, (Address r, int row, int col) cell) => DecodeAddress(ctx, GetAddress(cell));
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <returns></returns>
        public static string GetAddress(this IExcelContext ctx, Address r, int fromRow, string fromColumn, int toRow, string toColumn) => DecodeAddress(ctx, GetAddress(r, fromRow, fromColumn, toRow, toColumn));
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromColumn">From column.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toColumn">To column.</param>
        /// <returns></returns>
        public static string GetAddress(this IExcelContext ctx, Address r, int fromRow, int fromColumn, int toRow, int toColumn) => DecodeAddress(ctx, GetAddress(r, fromRow, fromColumn, toRow, toColumn));
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="cell">The cell.</param>
        /// <returns></returns>
        public static string GetAddress(this IExcelContext ctx, (Address r, int fromRow, int fromColumn, int toRow, int toColumn) cell) => DecodeAddress(ctx, GetAddress(cell));

        /// <summary>
        /// Decodes the address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="address">The address.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// address
        /// or
        /// address
        /// or
        /// address
        /// </exception>
        public static string DecodeAddress(this IExcelContext ctx, string address)
        {
            if (!address.StartsWith("^")) return address;
            var vec = address.Substring(1).Split(':').Select(x => int.Parse(x)).ToArray();
            var rel = (vec[0] & (int)Address.InternalRel) == (int)Address.InternalRel;
            if (vec.Length == 3)
            {
                int row = rel ? ctx.Y + vec[1] : vec[1],
                    col = rel ? ctx.X + vec[2] : vec[2],
                    coltocol1 = rel ? ctx.X + vec[1] : vec[1],
                    coltocol2 = rel ? ctx.X + vec[2] : vec[2],
                    rowtorow1 = rel ? ctx.Y + vec[1] : vec[1],
                    rowtorow2 = rel ? ctx.Y + vec[2] : vec[2];
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.CellAbs: return ExcelCellBase.GetAddress(row, col);
                    case Address.RangeAbs: return ExcelCellBase.GetAddress(ctx.Y, ctx.X, row, col);
                    case Address.RowOrCol: return row != 0 ? ExcelCellBase.GetAddressRow(row) : col != 0 ? ExcelCellBase.GetAddressCol(col) : null;
                    case Address.ColToCol: return $"{ExcelCellBase.GetAddressCol(coltocol1).Split(':')[0]}:{ExcelCellBase.GetAddressCol(coltocol2).Split(':')[0]}";
                    case Address.RowToRow: return $"{rowtorow1}:{rowtorow2}";
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            }
            else if (vec.Length == 5)
            {
                int fromRow = rel ? ctx.Y + vec[1] : vec[1], toRow = rel ? ctx.Y + vec[3] : vec[3],
                    fromCol = rel ? ctx.X + vec[2] : vec[2], toCol = rel ? ctx.X + vec[4] : vec[4];
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.RangeAbs: return ExcelCellBase.GetAddress(fromRow, fromCol, toRow, toCol);
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            }
            else throw new ArgumentOutOfRangeException(nameof(address));
        }

        internal static string DescribeAddress(string address)
        {
            if (!address.StartsWith("^")) return address;
            var vec = address.Substring(1).Split(':').Select(x => int.Parse(x)).ToArray();
            var rel = (vec[0] & (int)Address.InternalRel) == (int)Address.InternalRel ? "+" : null;
            if (vec.Length == 3)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.CellAbs: return $"r:{rel}{vec[1]}.{rel}{vec[2]}";
                    case Address.RangeAbs: return $"r:y.x:{rel}{vec[1]}.{rel}{vec[2]}";
                    case Address.RowOrCol: return $"r:r.c:{rel}{vec[1]}.{rel}{vec[2]}";
                    case Address.ColToCol: return $"r:c2c:{rel}{vec[1]}.{rel}{vec[2]}";
                    case Address.RowToRow: return $"r:r2r:{rel}{vec[1]}.{rel}{vec[2]}";
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            else if (vec.Length == 5)
                switch ((Address)(vec[0] & 0xF))
                {
                    case Address.RangeAbs: return $"r:{rel}{vec[1]}.{rel}{vec[2]}:{rel}{vec[3]}.{rel}{vec[4]}";
                    default: throw new ArgumentOutOfRangeException(nameof(address));
                }
            else throw new ArgumentOutOfRangeException(nameof(address));
        }

        internal static object ParseValue(this string v) =>
            v == null ? null :
                v.Contains("/") && DateTime.TryParse(v, out var vd) ? vd :
                v.Contains(".") && double.TryParse(v, out var vf) ? vf :
                long.TryParse(v, out var vl) ? vl :
                (object)v;

        /// <summary>
        /// Serializes the value.
        /// </summary>
        /// <param name="v">The v.</param>
        /// <param name="type">The type.</param>
        /// <returns></returns>
        internal static string SerializeValue(this object v, Type type) =>
            type == null ? v.ToString() : type.IsArray
            ? string.Join("|^", ((Array)v).Cast<object>().Select(x => x?.ToString()))
            : v.ToString();

        internal static object DeserializeValue(this string v, Type type) =>
            type == null ? v : type.IsArray && type.GetElementType() is Type elemType
            ? v.Split(new[] { "|^" }, StringSplitOptions.None).Select(x => Convert.ChangeType(x, elemType)).ToArray()
            : Convert.ChangeType(v, type);

        internal static T CastValue<T>(this object v, T defaultValue = default) => v != null
            ? typeof(T).IsArray && typeof(T).GetElementType() is Type elemType
            ? (T)(object)((Array)v).Cast<object>().Select(x => Convert.ChangeType(v, elemType)).ToArray()
            : (T)Convert.ChangeType(v, typeof(T))
            : defaultValue;

        internal static int RowToInt(string row) => !string.IsNullOrEmpty(row) ? int.Parse(row) : 0;
        internal static int ColToInt(string col) => col.ToUpperInvariant().Aggregate(0, (a, x) => (a * 26) + (x - '@'));
        internal static void CellToInts(string cell, out int row, out int col) { var idx = cell.IndexOfAny("0123456789".ToCharArray()); var val = idx != -1 ? new[] { cell.Substring(0, idx), cell.Substring(idx) } : new[] { cell, string.Empty }; row = RowToInt(val[1]); col = ColToInt(val[0]); }

        #region Measure

        /// <summary>
        /// Sets the width of the true column.
        /// </summary>
        /// <param name="column">The column.</param>
        /// <param name="width">The width.</param>
        public static void SetTrueColumnWidth(this ExcelColumn column, double width)
        {
            // Deduce what the column width would really get set to.
            var z = width >= (1 + 2 / 3)
                ? Math.Round((Math.Round(7 * (width - 1 / 256), 0) - 5) / 7, 2)
                : Math.Round((Math.Round(12 * (width - 1 / 256), 0) - Math.Round(5 * width, 0)) / 12, 2);

            // How far off? (will be less than 1)
            var errorAmt = width - z;

            // Calculate what amount to tack onto the original amount to result in the closest possible setting.
            var adj = width >= 1 + 2 / 3
                ? Math.Round(7 * errorAmt - 7 / 256, 0) / 7
                : Math.Round(12 * errorAmt - 12 / 256, 0) / 12 + (2 / 12);

            // Set width to a scaled-value that should result in the nearest possible value to the true desired setting.
            if (z > 0)
            {
                column.Width = width + adj;
                return;
            }
            column.Width = 0d;
        }

        // https://docs.microsoft.com/en-us/office/troubleshoot/excel/determine-column-widths#:~:text=In%20a%20new%20Excel%20workbook,)%2C%20and%20then%20click%20OK.
        //public static (int pixelWidth, int pixelHeight) GetAddressPixelSize(this IExcelContext ctx, ExcelAddressBase value)
        //{
        //    var wb = ((ExcelContext)ctx).WB;
        //    var normalFont = wb.Styles.NamedStyles.FirstOrDefault(x => x.Name == "Normal")?.Style.Font;

        //    var ws = ((ExcelContext)ctx).WS;
        //    double pixelWidth = 0f, pixelHeight = 0f;
        //    for (var col = value.Start.Column; col <= value.End.Column; col++) pixelWidth += ws.Column(col).Width;
        //    for (var row = value.Start.Row; row <= value.End.Row; row++) pixelHeight += ws.Row(row).Height;
        //    return ((int)pixelWidth, (int)pixelHeight);
        //}

        #endregion

        #region Lookup

        // reference extracted from ECMA-376, Part 4, Section 3.8.26 or 18.8.27 SE Part 1
        internal readonly static Color[] ExcelColorLookup = new[] {
            ColorTranslator.FromHtml("#000000"), // 0
            ColorTranslator.FromHtml("#FFFFFF"),
            ColorTranslator.FromHtml("#FF0000"),
            ColorTranslator.FromHtml("#00FF00"),
            ColorTranslator.FromHtml("#0000FF"),
            ColorTranslator.FromHtml("#FFFF00"),
            ColorTranslator.FromHtml("#FF00FF"),
            ColorTranslator.FromHtml("#00FFFF"),
            ColorTranslator.FromHtml("#000000"), // 8
            ColorTranslator.FromHtml("#FFFFFF"),
            ColorTranslator.FromHtml("#FF0000"),
            ColorTranslator.FromHtml("#00FF00"),
            ColorTranslator.FromHtml("#0000FF"),
            ColorTranslator.FromHtml("#FFFF00"),
            ColorTranslator.FromHtml("#FF00FF"),
            ColorTranslator.FromHtml("#00FFFF"),
            ColorTranslator.FromHtml("#800000"),
            ColorTranslator.FromHtml("#008000"),
            ColorTranslator.FromHtml("#000080"),
            ColorTranslator.FromHtml("#808000"),
            ColorTranslator.FromHtml("#800080"),
            ColorTranslator.FromHtml("#008080"),
            ColorTranslator.FromHtml("#C0C0C0"),
            ColorTranslator.FromHtml("#808080"),
            ColorTranslator.FromHtml("#9999FF"),
            ColorTranslator.FromHtml("#993366"),
            ColorTranslator.FromHtml("#FFFFCC"),
            ColorTranslator.FromHtml("#CCFFFF"),
            ColorTranslator.FromHtml("#660066"),
            ColorTranslator.FromHtml("#FF8080"),
            ColorTranslator.FromHtml("#0066CC"),
            ColorTranslator.FromHtml("#CCCCFF"),
            ColorTranslator.FromHtml("#000080"),
            ColorTranslator.FromHtml("#FF00FF"),
            ColorTranslator.FromHtml("#FFFF00"),
            ColorTranslator.FromHtml("#00FFFF"),
            ColorTranslator.FromHtml("#800080"),
            ColorTranslator.FromHtml("#800000"),
            ColorTranslator.FromHtml("#008080"),
            ColorTranslator.FromHtml("#0000FF"),
            ColorTranslator.FromHtml("#00CCFF"),
            ColorTranslator.FromHtml("#CCFFFF"),
            ColorTranslator.FromHtml("#CCFFCC"),
            ColorTranslator.FromHtml("#FFFF99"),
            ColorTranslator.FromHtml("#99CCFF"),
            ColorTranslator.FromHtml("#FF99CC"),
            ColorTranslator.FromHtml("#CC99FF"),
            ColorTranslator.FromHtml("#FFCC99"),
            ColorTranslator.FromHtml("#3366FF"),
            ColorTranslator.FromHtml("#33CCCC"),
            ColorTranslator.FromHtml("#99CC00"),
            ColorTranslator.FromHtml("#FFCC00"),
            ColorTranslator.FromHtml("#FF9900"),
            ColorTranslator.FromHtml("#FF6600"),
            ColorTranslator.FromHtml("#666699"),
            ColorTranslator.FromHtml("#969696"),
            ColorTranslator.FromHtml("#003366"),
            ColorTranslator.FromHtml("#339966"),
            ColorTranslator.FromHtml("#003300"),
            ColorTranslator.FromHtml("#333300"),
            ColorTranslator.FromHtml("#993300"),
            ColorTranslator.FromHtml("#993366"),
            ColorTranslator.FromHtml("#333399"),
            ColorTranslator.FromHtml("#333333"), // 63
        };

        #endregion
    }
}
