using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Name` range of `.NameKind` to the `.Cells` in range
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct WorkbookName : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; private set; }
        /// <summary>
        /// Gets the cells.
        /// </summary>
        /// <value>
        /// The cells.
        /// </value>
        public string Cells { get; private set; }
        /// <summary>
        /// Gets the kind of the name.
        /// </summary>
        /// <value>
        /// The kind of the name.
        /// </value>
        public WorkbookNameKind NameKind { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookName"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public WorkbookName(string name, int row, int col, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(row, col), nameKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookName"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public WorkbookName(string name, int fromRow, int fromCol, int toRow, int toCol, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), nameKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookName"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="r">The r.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public WorkbookName(string name, Address r, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(r, 0, 0), nameKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookName"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public WorkbookName(string name, Address r, int row, int col, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(r, row, col), nameKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookName"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public WorkbookName(string name, Address r, int fromRow, int fromCol, int toRow, int toCol, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), nameKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookName"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="cells">The cells.</param>
        /// <param name="nameKind">Kind of the name.</param>
        /// <exception cref="ArgumentNullException">
        /// cells
        /// or
        /// cells
        /// </exception>
        public WorkbookName(string name, string cells, WorkbookNameKind nameKind = WorkbookNameKind.Add)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(cells));
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            NameKind = nameKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
            Cells = r.ReadString();
            NameKind = (WorkbookNameKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
            w.Write(Cells);
            w.Write((int)NameKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorkbookName(Name, Cells, NameKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorkbookName[{Name}]: {ExcelService.DescribeAddress(Cells)} - {NameKind}"); }
    }
}