using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Styles` to the `.Cells` in range
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct CellStyle : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the cells.
        /// </summary>
        /// <value>
        /// The cells.
        /// </value>
        public string Cells { get; private set; }
        /// <summary>
        /// Gets the styles.
        /// </summary>
        /// <value>
        /// The styles.
        /// </value>
        public string[] Styles { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CellStyle"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="styles">The styles.</param>
        public CellStyle(int row, int col, params string[] styles)
            : this(ExcelService.GetAddress(row, col), styles) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellStyle"/> struct.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="styles">The styles.</param>
        public CellStyle(int fromRow, int fromCol, int toRow, int toCol, params string[] styles)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), styles) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellStyle"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="styles">The styles.</param>
        public CellStyle(Address r, params string[] styles)
            : this(ExcelService.GetAddress(r, 0, 0), styles) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellStyle"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="styles">The styles.</param>
        public CellStyle(Address r, int row, int col, params string[] styles)
            : this(ExcelService.GetAddress(r, row, col), styles) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellStyle"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="styles">The styles.</param>
        public CellStyle(Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] styles)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), styles) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellStyle"/> struct.
        /// </summary>
        /// <param name="cells">The cells.</param>
        /// <param name="styles">The styles.</param>
        /// <exception cref="ArgumentNullException">
        /// cells
        /// or
        /// styles
        /// </exception>
        public CellStyle(string cells, params string[] styles)
        {
            When = When.Normal;
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            Styles = styles ?? throw new ArgumentNullException(nameof(styles));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Styles = new string[r.ReadUInt16()];
            for (var i = 0; i < Styles.Length; i++)
                Styles[i] = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write((ushort)Styles.Length);
            for (var i = 0; i < Styles.Length; i++)
                w.Write(Styles[i]);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CellStyle(Cells, Styles);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CellStyle[{ExcelService.DescribeAddress(Cells)}]: {string.Join(", ", Styles)}"); }
    }
}