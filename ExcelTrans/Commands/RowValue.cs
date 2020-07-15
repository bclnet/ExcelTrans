using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Value` of `.ValueKind` to the `.Row` row
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct RowValue : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the row.
        /// </summary>
        /// <value>
        /// The row.
        /// </value>
        public int Row { get; private set; }
        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public string Value { get; private set; }
        /// <summary>
        /// Gets the kind of the value.
        /// </summary>
        /// <value>
        /// The kind of the value.
        /// </value>
        public RowValueKind ValueKind { get; private set; }
        /// <summary>
        /// Gets or sets the type of the value.
        /// </summary>
        /// <value>
        /// The type of the value.
        /// </value>
        public Type ValueType { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="RowValue"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public RowValue(string row, object value, RowValueKind valueKind) : this(ExcelService.RowToInt(row), value, valueKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="RowValue"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public RowValue(int row, object value, RowValueKind valueKind)
        {
            When = When.Normal;
            Row = row;
            ValueKind = valueKind;
            ValueType = value?.GetType();
            Value = value?.SerializeValue(ValueType);
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Row = r.ReadInt32();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (RowValueKind)r.ReadInt32();
            ValueType = r.ReadBoolean() ? Type.GetType(r.ReadString()) : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Row);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
            w.Write(ValueType != null); if (ValueType != null) w.Write(ValueType.ToString());
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.RowValue(Row, Value?.DeserializeValue(ValueType), ValueKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}RowValue[{Row}]: {Value}{$" - {ValueKind}"}"); }

    }
}