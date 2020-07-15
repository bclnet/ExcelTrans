using System;
using System.IO;
using System.Text.Json;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies json `.Value` of `.FormattingKind` to `.Address` in range
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct ConditionalFormatting : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the address.
        /// </summary>
        /// <value>
        /// The address.
        /// </value>
        public string Address { get; private set; }
        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public string Value { get; private set; }
        /// <summary>
        /// Gets the kind of the formatting.
        /// </summary>
        /// <value>
        /// The kind of the formatting.
        /// </value>
        public ConditionalFormattingKind FormattingKind { get; private set; }
        /// <summary>
        /// Gets the priority.
        /// </summary>
        /// <value>
        /// The priority.
        /// </value>
        public int? Priority { get; private set; }
        /// <summary>
        /// Gets a value indicating whether [stop if true].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [stop if true]; otherwise, <c>false</c>.
        /// </value>
        public bool StopIfTrue { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ConditionalFormatting"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public ConditionalFormatting(int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false)
            : this(ExcelService.GetAddress(row, col), value, formattingKind, priority, stopIfTrue) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ConditionalFormatting"/> struct.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public ConditionalFormatting(int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ConditionalFormatting"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public ConditionalFormatting(Address r, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false)
            : this(ExcelService.GetAddress(r, 0, 0), value, formattingKind, priority, stopIfTrue) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ConditionalFormatting"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public ConditionalFormatting(Address r, int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false)
            : this(ExcelService.GetAddress(r, row, col), value, formattingKind, priority, stopIfTrue) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ConditionalFormatting"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public ConditionalFormatting(Address r, int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ConditionalFormatting"/> struct.
        /// </summary>
        /// <param name="address">The address.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        /// <exception cref="ArgumentNullException">address</exception>
        public ConditionalFormatting(string address, object value, ConditionalFormattingKind formattingKind, int? priority = null, bool stopIfTrue = false)
        {
            When = When.Normal;
            Address = address ?? throw new ArgumentNullException(nameof(address));
            Value = value != null ? value is string @string ? @string : JsonSerializer.Serialize(value) : null;
            FormattingKind = formattingKind;
            Priority = priority;
            StopIfTrue = stopIfTrue;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Address = r.ReadString();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            FormattingKind = (ConditionalFormattingKind)r.ReadInt32();
            Priority = r.ReadBoolean() ? (int?)r.ReadInt32() : null;
            StopIfTrue = r.ReadBoolean();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Address);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)FormattingKind);
            w.Write(Priority != null); if (Priority != null) w.Write(Priority.Value);
            w.Write(StopIfTrue);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.ConditionalFormatting(Address, Value, FormattingKind, Priority, StopIfTrue);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}ConditionalFormatting[{Address}]: {Value} - {FormattingKind}"); }
    }
}