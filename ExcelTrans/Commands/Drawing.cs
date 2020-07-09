using System;
using System.IO;
using System.Text.Json;

namespace ExcelTrans.Commands
{
    public struct Drawing : IExcelCommand
    {
        public When When { get; }
        public string Address { get; private set; }
        public string Name { get; private set; }
        public string Value { get; private set; }
        public DrawingKind DrawingKind { get; private set; }

        public Drawing(int row, int col, string name, object value, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(row, col), name, value, drawingKind) { }
        public Drawing(int fromRow, int fromCol, int toRow, int toCol, string name, object value, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), name, value, drawingKind) { }
        public Drawing(Address r, string name, object value, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(r, 0, 0), name, value, drawingKind) { }
        public Drawing(Address r, int row, int col, string name, object value, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(r, row, col), name, value, drawingKind) { }
        public Drawing(Address r, int fromRow, int fromCol, int toRow, int toCol, string name, object value, DrawingKind drawingKind)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), name, value, drawingKind) { }
        public Drawing(string address, string name, object value, DrawingKind drawingKind)
        {
            When = When.Normal;
            Address = address;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Value = value != null ? value is string @string ? @string : JsonSerializer.Serialize(value) : null;
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