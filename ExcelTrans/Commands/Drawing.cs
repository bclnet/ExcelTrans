using System;
using System.IO;
using System.Text.Json;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies json `.Value` of `.DrawingKind` with `.Name` to `.Address` in range
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct Drawing : IExcelCommand
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
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; private set; }
        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public string Value { get; private set; }
        /// <summary>
        /// Gets the kind of the drawing.
        /// </summary>
        /// <value>
        /// The kind of the drawing.
        /// </value>
        public DrawingKind DrawingKind { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Drawing"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="name">The name.</param>
        /// <param name="json">The json.</param>
        /// <param name="drawingKind">Kind of the drawing.</param>
        public Drawing(int row, int col, string name, object json, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(row, col), name, json, drawingKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Drawing"/> struct.
        /// </summary>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="name">The name.</param>
        /// <param name="json">The json.</param>
        /// <param name="drawingKind">Kind of the drawing.</param>
        public Drawing(int fromRow, int fromCol, int toRow, int toCol, string name, object json, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), name, json, drawingKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Drawing"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="name">The name.</param>
        /// <param name="json">The json.</param>
        /// <param name="drawingKind">Kind of the drawing.</param>
        public Drawing(Address r, string name, object json, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(r, 0, 0), name, json, drawingKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Drawing"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="name">The name.</param>
        /// <param name="json">The json.</param>
        /// <param name="drawingKind">Kind of the drawing.</param>
        public Drawing(Address r, int row, int col, string name, object json, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(r, row, col), name, json, drawingKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Drawing"/> struct.
        /// </summary>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="name">The name.</param>
        /// <param name="json">The json.</param>
        /// <param name="drawingKind">Kind of the drawing.</param>
        public Drawing(Address r, int fromRow, int fromCol, int toRow, int toCol, string name, object json, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), name, json, drawingKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Drawing"/> struct.
        /// </summary>
        /// <param name="address">The address.</param>
        /// <param name="name">The name.</param>
        /// <param name="json">The json.</param>
        /// <param name="drawingKind">Kind of the drawing.</param>
        /// <exception cref="ArgumentNullException">name</exception>
        public Drawing(string address, string name, object json, DrawingKind drawingKind)
        {
            When = When.Normal;
            Address = address;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Value = json != null ? json is string @string ? @string : JsonSerializer.Serialize(json) : null;
            DrawingKind = drawingKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Address = r.ReadBoolean() ? r.ReadString() : null;
            Name = r.ReadString();
            Value = r.ReadBoolean() ? r.ReadString() : null;
            DrawingKind = (DrawingKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Address != null); if (Address != null) w.Write(Address);
            w.Write(Name);
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)DrawingKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Drawing(Address, Name, Value, DrawingKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}Drawing[{Address}]: {Name} - {DrawingKind}"); }
    }
}