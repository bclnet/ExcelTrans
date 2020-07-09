using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct Protection : IExcelCommand
    {
        public When When { get; }
        public string Value { get; private set; }
        public WorkbookProtectionKind ProtectionKind { get; private set; }

        public Protection(int row, int col, WorkbookProtectionKind protectionKind)
            : this(ExcelService.GetAddress(row, col), protectionKind) { }
        public Protection(string value, WorkbookProtectionKind protectionKind)
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

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Protection(Value, ProtectionKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}Protection: {Value} - {ProtectionKind}"); }
    }
}