using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Pops a Set off the context stack
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct PopSet : IExcelCommand
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

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Sets.Pop().Execute(ctx);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PopSet"); }

        internal static void Flush(IExcelContext ctx, int index)
        {
            while (ctx.Sets.Count > index)
                ctx.Sets.Pop().Execute(ctx);
        }

    }
}