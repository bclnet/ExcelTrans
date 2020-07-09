using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct VbaCodeModule : IExcelCommand
    {
        public When When { get; }
        public string Name { get; private set; }
        public VbaCode Code { get; private set; }
        public VbaModuleKind ModuleKind { get; private set; }

        public VbaCodeModule(VbaCode code) : this("CodeModule", code, VbaModuleKind.CodeModule) { }
        public VbaCodeModule(string name, VbaCode code, VbaModuleKind actionKind = VbaModuleKind.Module)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Code = code ?? throw new ArgumentNullException(nameof(code));
            ModuleKind = actionKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
            Code = new VbaCode().Read(r);
            ModuleKind = (VbaModuleKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
            Code.Write(w);
            w.Write((int)ModuleKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.VbaCodeModule(Name, Code, ModuleKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}VbaCodeModule: {Name} - {ModuleKind}"); }
    }
}