using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Value` of `.ProtectionKind` to Workbook
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct WorkbookProtection : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public string Value { get; private set; }
        /// <summary>
        /// Gets the kind of the protection.
        /// </summary>
        /// <value>
        /// The kind of the protection.
        /// </value>
        public WorkbookProtectionKind ProtectionKind { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookProtection"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="protectionKind">Kind of the protection.</param>
        public WorkbookProtection(int row, int col, WorkbookProtectionKind protectionKind)
            : this(ExcelService.GetAddress(row, col), protectionKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookProtection"/> struct.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="protectionKind">Kind of the protection.</param>
        public WorkbookProtection(string value, WorkbookProtectionKind protectionKind)
        {
            When = When.Normal;
            Value = value;
            ProtectionKind = protectionKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ProtectionKind = (WorkbookProtectionKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ProtectionKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorkbookProtection(Value, ProtectionKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorkbookProtection: {Value} - {ProtectionKind}"); }
    }
}