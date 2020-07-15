using System;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Adds `.Libraries` of type VbaLibrary to the VbaProject
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct VbaReference : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the libraries.
        /// </summary>
        /// <value>
        /// The libraries.
        /// </value>
        public VbaLibrary[] Libraries { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VbaReference"/> struct.
        /// </summary>
        /// <param name="libraries">The libraries.</param>
        /// <exception cref="ArgumentNullException">libraries</exception>
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