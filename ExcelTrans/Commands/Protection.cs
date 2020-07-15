using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Value` of `.ProtectionKind` to current worksheet
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct Protection : IExcelCommand
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
        public ProtectionKind ProtectionKind { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Protection"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="protectionKind">Kind of the protection.</param>
        public Protection(int row, int col, ProtectionKind protectionKind)
            : this(ExcelService.GetAddress(row, col), protectionKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Protection"/> struct.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="protectionKind">Kind of the protection.</param>
        public Protection(string value, ProtectionKind protectionKind)
        {
            When = When.Normal;
            Value = value;
            ProtectionKind = protectionKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ProtectionKind = (ProtectionKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ProtectionKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Protection(Value, ProtectionKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}Protection: {Value} - {ProtectionKind}"); }
    }
}