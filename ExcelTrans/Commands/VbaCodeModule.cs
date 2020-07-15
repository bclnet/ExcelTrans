using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Code` of `.ModuleKind` with `.Name` to the VbaProject
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct VbaCodeModule : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; private set; }
        /// <summary>
        /// Gets the code.
        /// </summary>
        /// <value>
        /// The code.
        /// </value>
        public VbaCode Code { get; private set; }
        /// <summary>
        /// Gets the kind of the module.
        /// </summary>
        /// <value>
        /// The kind of the module.
        /// </value>
        public VbaModuleKind ModuleKind { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCodeModule"/> struct.
        /// </summary>
        /// <param name="code">The code.</param>
        public VbaCodeModule(VbaCode code) : this("CodeModule", code, VbaModuleKind.CodeModule) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCodeModule"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="code">The code.</param>
        /// <param name="actionKind">Kind of the action.</param>
        /// <exception cref="ArgumentNullException">
        /// name
        /// or
        /// code
        /// </exception>
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