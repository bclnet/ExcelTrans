using ExcelTrans.Utils;
using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Executes `.Action()`
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public class Command : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; private set; }
        /// <summary>
        /// Gets the action.
        /// </summary>
        /// <value>
        /// The action.
        /// </value>
        public Action<IExcelContext> Action { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// </summary>
        /// <param name="action">The action.</param>
        public Command(Action action)
            : this(When.Normal, v => action()) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// </summary>
        /// <param name="action">The action.</param>
        public Command(Action<IExcelContext> action)
            : this(When.Normal, action) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="action">The action.</param>
        public Command(When when, Action action)
            : this(when, v => action()) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="action">The action.</param>
        /// <exception cref="ArgumentNullException">action</exception>
        public Command(When when, Action<IExcelContext> action)
        {
            When = when;
            Action = action ?? throw new ArgumentNullException(nameof(action));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            When = (When)r.ReadByte();
            Action = ExcelSerDes.DecodeAction<IExcelContext>(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write((byte)When);
            ExcelSerDes.EncodeAction(w, Action);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => Action(ctx);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}Command{(When == When.Normal ? null : $"[{When}]")}: [action]"); }
    }
}