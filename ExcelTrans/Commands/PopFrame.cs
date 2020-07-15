using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Pops a Frame off the context stack
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct PopFrame : IExcelCommand
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

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Frame = ctx.Frames.Pop();

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PopFrame"); }
    }
}