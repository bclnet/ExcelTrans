using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Value` of `.ValueKind` to the `.Col` column
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct ColumnValue : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the col.
        /// </summary>
        /// <value>
        /// The col.
        /// </value>
        public int Col { get; private set; }
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
        public ColumnValueKind ValueKind { get; private set; }
        /// <summary>
        /// Gets or sets the type of the value.
        /// </summary>
        /// <value>
        /// The type of the value.
        /// </value>
        public Type ValueType { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnValue"/> struct.
        /// </summary>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public ColumnValue(string col, object value, ColumnValueKind valueKind) : this(ExcelService.ColToInt(col), value, valueKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ColumnValue"/> struct.
        /// </summary>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public ColumnValue(int col, object value, ColumnValueKind valueKind)
        {
            When = When.Normal;
            Col = col;
            ValueKind = valueKind;
            ValueType = value?.GetType();
            Value = value?.SerializeValue(ValueType);
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Col = r.ReadInt32();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (ColumnValueKind)r.ReadInt32();
            ValueType = r.ReadBoolean() ? Type.GetType(r.ReadString()) : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Col);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
            w.Write(ValueType != null); if (ValueType != null) w.Write(ValueType.ToString());
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.ColumnValue(Col, Value?.DeserializeValue(ValueType), ValueKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}ColumnValue[{Col}]: {Value}{$" - {ValueKind}"}"); }
    }
}