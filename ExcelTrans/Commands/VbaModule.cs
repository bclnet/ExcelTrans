using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Code` with `.Name` to the VbaProject
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
        /// Initializes a new instance of the <see cref="VbaModule"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="code">The code.</param>
        /// <exception cref="ArgumentNullException">
        /// code
        /// or
        /// code
        /// </exception>
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