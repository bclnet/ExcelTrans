using System;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    public struct VbaReference : IExcelCommand
    {
        public When When { get; }
        public VbaLibrary[] Libraries { get; private set; }

        public VbaReference(params VbaLibrary[] libraries)
        {
            When = When.Normal;
            Libraries = libraries ?? throw new ArgumentNullException(nameof(libraries));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Libraries = new VbaLibrary[r.ReadUInt16()];
            for (var i = 0; i < Libraries.Length; i++)
                Libraries[i] = new VbaLibrary().Read(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write((ushort)Libraries.Length);
            for (var i = 0; i < Libraries.Length; i++)
                Libraries[i].Write(w);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.VbaReference(Libraries);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}VbaReference: {string.Join(", ", Libraries.Select(x => x.Name))}"); }
    }
}