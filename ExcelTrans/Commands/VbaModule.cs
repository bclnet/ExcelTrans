using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct VbaModule : IExcelCommand
    {
        public When When { get; }
        public string Name { get; private set; }
        public VbaCode Code { get; private set; }

        public VbaModule(string name, VbaCode code)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(code));
            Code = code ?? throw new ArgumentNullException(nameof(code));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
            Code = new VbaCode().Read(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
            Code.Write(w);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.VbaModule(Name, Code);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}VbaModule: {Name} - {Code.Name ?? "{default}"}"); }
    }
}