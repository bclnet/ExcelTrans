using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Rules` of `.ValidationKind` to the `.Cells` in range
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct CellValidation : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the kind of the validation.
        /// </summary>
        /// <value>
        /// The kind of the validation.
        /// </value>
        public CellValidationKind ValidationKind { get; private set; }
        /// <summary>
        /// Gets the cells.
        /// </summary>
        /// <value>
        /// The cells.
        /// </value>
        public string Cells { get; private set; }
        /// <summary>
        /// Gets the rules.
        /// </summary>
        /// <value>
        /// The rules.
        /// </value>
        public string[] Rules { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CellValidation"/> struct.
        /// </summary>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="rules">The rules.</param>
        public CellValidation(CellValidationKind validationKind, int row, int col, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(row, col), rules) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValidation"/> struct.
        /// </summary>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="rules">The rules.</param>
        public CellValidation(CellValidationKind validationKind, int fromRow, int fromCol, int toRow, int toCol, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), rules) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValidation"/> struct.
        /// </summary>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="r">The r.</param>
        /// <param name="rules">The rules.</param>
        public CellValidation(CellValidationKind validationKind, Address r, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(r, 0, 0), rules) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValidation"/> struct.
        /// </summary>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="rules">The rules.</param>
        public CellValidation(CellValidationKind validationKind, Address r, int row, int col, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(r, row, col), rules) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValidation"/> struct.
        /// </summary>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="rules">The rules.</param>
        public CellValidation(CellValidationKind validationKind, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), rules) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValidation"/> struct.
        /// </summary>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="cells">The cells.</param>
        /// <param name="rules">The rules.</param>
        /// <exception cref="ArgumentNullException">
        /// cells
        /// or
        /// rules
        /// </exception>
        public CellValidation(CellValidationKind validationKind, string cells, params string[] rules)
        {
            When = When.Normal;
            ValidationKind = validationKind;
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            Rules = rules ?? throw new ArgumentNullException(nameof(rules));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            ValidationKind = (CellValidationKind)r.ReadInt32();
            Rules = new string[r.ReadUInt16()];
            for (var i = 0; i < Rules.Length; i++)
                Rules[i] = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write((int)ValidationKind);
            w.Write((ushort)Rules.Length);
            for (var i = 0; i < Rules.Length; i++)
                w.Write(Rules[i]);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CellValidation(ValidationKind, Cells, Rules);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CellValidation.{ValidationKind}[{ExcelService.DescribeAddress(Cells)}]: {string.Join(", ", Rules)}"); }
    }
}