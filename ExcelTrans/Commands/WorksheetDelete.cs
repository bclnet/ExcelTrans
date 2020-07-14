using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct WorksheetDelete : IExcelCommand
    {
        public When When { get; }
        public string Name { get; private set; }

        public WorksheetDelete(string name)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        void IExcelCommand.Read(BinaryReader r) => Name = r.ReadString();

        void IExcelCommand.Write(BinaryWriter w) => w.Write(Name);

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorksheetDelete(Name);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorksheetDelete: {Name}"); }
    }
}