using ExcelTrans.Utils;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Executes `.Func()` per Row
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public class CommandRow : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; private set; }
        /// <summary>
        /// Gets the function.
        /// </summary>
        /// <value>
        /// The function.
        /// </value>
        public Func<IExcelContext, Collection<string>, CommandRtn> Func { get; private set; }
        /// <summary>
        /// Gets the CMDS.
        /// </summary>
        /// <value>
        /// The CMDS.
        /// </value>
        public IExcelCommand[] Cmds { get; private set; }

        // action - iexcelcommand[]
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(Action func, params IExcelCommand[] cmds)
            : this(When.Normal, (a, b) => func(), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(Action<IExcelContext, Collection<string>> func, params IExcelCommand[] cmds)
            : this(When.Normal, func, cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(When when, Action func, params IExcelCommand[] cmds)
            : this(when, (a, b) => func(), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Action<IExcelContext, Collection<string>> func, params IExcelCommand[] cmds)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // action - action
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Action func, Action command)
            : this(When.Normal, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Action<IExcelContext, Collection<string>> func, Action command)
            : this(When.Normal, func, command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(When when, Action func, Action command)
            : this(when, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Action<IExcelContext, Collection<string>> func, Action command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // action - action<iexcelcontext>
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Action func, Action<IExcelContext> command)
            : this(When.Normal, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Action<IExcelContext, Collection<string>> func, Action<IExcelContext> command)
            : this(When.Normal, func, command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(When when, Action func, Action<IExcelContext> command)
            : this(when, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Action<IExcelContext, Collection<string>> func, Action<IExcelContext> command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<bool> - iexcelcommand[]
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(Func<bool> func, params IExcelCommand[] cmds)
            : this(When.Normal, (a, b) => func(), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(Func<IExcelContext, Collection<string>, bool> func, params IExcelCommand[] cmds)
            : this(When.Normal, func, cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(When when, Func<bool> func, params IExcelCommand[] cmds)
            : this(when, (a, b) => func(), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Func<IExcelContext, Collection<string>, bool> func, params IExcelCommand[] cmds)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<bool> - action
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<bool> func, Action command)
            : this(When.Normal, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<IExcelContext, Collection<string>, bool> func, Action command)
            : this(When.Normal, func, command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(When when, Func<bool> func, Action command)
            : this(when, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Func<IExcelContext, Collection<string>, bool> func, Action command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<bool> - action<iexcelcontext>
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<bool> func, Action<IExcelContext> command)
            : this(When.Normal, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow" /> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<IExcelContext, Collection<string>, bool> func, Action<IExcelContext> command)
            : this(When.Normal, func, command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(When when, Func<bool> func, Action<IExcelContext> command)
            : this(when, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Func<IExcelContext, Collection<string>, bool> func, Action<IExcelContext> command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<commandrtn> - iexcelcommand[]
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(Func<CommandRtn> func, params IExcelCommand[] cmds)
            : this(When.Normal, (a, b) => func(), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(Func<IExcelContext, Collection<string>, CommandRtn> func, params IExcelCommand[] cmds)
            : this(When.Normal, func, cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandRow(When when, Func<CommandRtn> func, params IExcelCommand[] cmds)
            : this(when, (a, b) => func(), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Func<IExcelContext, Collection<string>, CommandRtn> func, params IExcelCommand[] cmds)
        {
            When = when;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<commandrtn> - action
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<CommandRtn> func, Action command)
            : this(When.Normal, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<IExcelContext, Collection<string>, CommandRtn> func, Action command)
            : this(When.Normal, func, command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(When when, Func<CommandRtn> func, Action command)
            : this(when, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Func<IExcelContext, Collection<string>, CommandRtn> func, Action command)
        {
            When = when;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<commandrtn> - action<iexcelcontext>
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<CommandRtn> func, Action<IExcelContext> command)
            : this(When.Normal, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(Func<IExcelContext, Collection<string>, CommandRtn> func, Action<IExcelContext> command)
            : this(When.Normal, func, command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandRow(When when, Func<CommandRtn> func, Action<IExcelContext> command)
            : this(when, (a, b) => func(), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandRow"/> class.
        /// </summary>
        /// <param name="when">The when.</param>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandRow(When when, Func<IExcelContext, Collection<string>, CommandRtn> func, Action<IExcelContext> command)
        {
            When = when;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            When = (When)r.ReadByte();
            Func = ExcelSerDes.DecodeFunc<IExcelContext, Collection<string>, CommandRtn>(r);
            Cmds = ExcelSerDes.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write((byte)When);
            ExcelSerDes.EncodeFunc(w, Func);
            ExcelSerDes.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CmdRows.Push(this);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CommandRow{(When == When.Normal ? null : $"[{When}]")}: [func]"); ExcelSerDes.DescribeCommands(w, pad, Cmds); }

        static Func<IExcelContext, Collection<string>, CommandRtn> FuncWrapper(Action<IExcelContext, Collection<string>> action) => (z, c) => { action(z, c); return CommandRtn.Normal; };

        static Func<IExcelContext, Collection<string>, CommandRtn> FuncWrapper(Func<IExcelContext, Collection<string>, bool> action) => (z, c) => action(z, c) ? CommandRtn.Normal : CommandRtn.SkipCmds;

        internal static void Flush(IExcelContext ctx, int idx)
        {
            while (ctx.CmdRows.Count > idx)
                ctx.CmdRows.Pop();
        }
    }
}