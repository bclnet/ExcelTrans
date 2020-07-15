using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Flushes all pending commands
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct Flush : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }

        void IExcelCommand.Read(BinaryReader r) { }

        void IExcelCommand.Write(BinaryWriter w) { }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Flush();

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}Flush"); }
    }
}