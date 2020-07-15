using ExcelTrans.Utils;
using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Pushes a new Frame with `cmds` onto the context stack
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct PushFrame : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the CMDS.
        /// </summary>
        /// <value>
        /// The CMDS.
        /// </value>
        public IExcelCommand[] Cmds { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="PushFrame"/> struct.
        /// </summary>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">cmds</exception>
        public PushFrame(params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Cmds = cmds ?? throw new ArgumentNullException(nameof(cmds));
        }

        void IExcelCommand.Read(BinaryReader r) => Cmds = ExcelSerDes.DecodeCommands(r);

        void IExcelCommand.Write(BinaryWriter w) => ExcelSerDes.EncodeCommands(w, Cmds);

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after)
        {
            ctx.Frames.Push(ctx.Frame);
            ctx.ExecuteCmd(Cmds, out after); //action?.Invoke();
        }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PushFrame:"); ExcelSerDes.DescribeCommands(w, pad, Cmds); }
    }
}