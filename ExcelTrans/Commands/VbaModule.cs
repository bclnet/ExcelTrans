using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Code` of `.ModuleKind` to the VbaProject
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct VbaModule : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
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
        /// Initializes a new instance of the <see cref="VbaModule"/> struct.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <param name="actionKind">Kind of the action.</param>
        /// <exception cref="ArgumentNullException">code</exception>
        public VbaModule(VbaCode code, VbaModuleKind actionKind = VbaModuleKind.Get)
        {
            When = When.Normal;
            Code = code ?? throw new ArgumentNullException(nameof(code));
            ModuleKind = actionKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Code = new VbaCode().Read(r);
            ModuleKind = (VbaModuleKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            Code.Write(w);
            w.Write((int)ModuleKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.VbaModule(Code, ModuleKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}VbaModule: {Code.Name} - {ModuleKind}"); }
    }
}