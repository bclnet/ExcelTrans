﻿using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Value` of `.ValueKind` to the `.Cells` in range
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct CellValue : IExcelCommand
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
        /// Gets or sets the type of the value.
        /// </summary>
        /// <value>
        /// The type of the value.
        /// </value>
        public Type ValueType { get; set; }
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
        public CellValueKind ValueKind { get; private set; }
        /// <summary>
        /// Gets the value format.
        /// </summary>
        /// <value>
        /// The value format.
        /// </value>
        public string ValueFormat { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CellValue" /> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CellValue(int row, int col, object value, CellValueKind valueKind = CellValueKind.Value, string valueFormat = null)
            : this(ExcelService.GetAddress(row, col), value, valueKind, valueFormat) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValue" /> struct.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CellValue(int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value, string valueFormat = null)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, valueKind, valueFormat) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValue" /> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="value">The value.</param>
        public CellValue(Address r, object value)
            : this(ExcelService.GetAddress(r, 0, 0), value, CellValueKind.Value, null) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValue" /> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CellValue(Address r, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value, string valueFormat = null)
            : this(ExcelService.GetAddress(r, row, col), value, valueKind, valueFormat) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValue" /> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CellValue(Address r, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value, string valueFormat = null)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, valueKind, valueFormat) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CellValue" /> struct.
        /// </summary>
        /// <param name="cells">The cells.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        /// <exception cref="ArgumentNullException">cells</exception>
        public CellValue(string cells, object value, CellValueKind valueKind = CellValueKind.Value, string valueFormat = null)
        {
            When = When.Normal;
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            ValueType = value?.GetType();
            Value = value?.SerializeValue(ValueType);
            ValueKind = valueKind;
            ValueFormat = valueFormat;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            ValueType = r.ReadBoolean() ? Type.GetType(r.ReadString()) : null;
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ValueKind = (CellValueKind)r.ReadInt32();
            ValueFormat = r.ReadBoolean() ? r.ReadString() : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write(ValueType != null); if (ValueType != null) w.Write(ValueType.ToString());
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ValueKind);
            w.Write(ValueFormat != null); if (ValueFormat != null) w.Write(ValueFormat);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CellValue(Cells, Value?.DeserializeValue(ValueType), ValueKind, ValueFormat);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CellValue[{ExcelService.DescribeAddress(Cells)}]: {Value}{(ValueKind == CellValueKind.Value ? null : $" - {ValueKind}")}"); }
    }
}