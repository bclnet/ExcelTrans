using ExcelTrans.Utils;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Executes `.Func()` per Column
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public class CommandCol : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the function.
        /// </summary>
        /// <value>
        /// The function.
        /// </value>
        public Func<IExcelContext, Collection<string>, object, CommandRtn> Func { get; private set; }
        /// <summary>
        /// Gets the CMDS.
        /// </summary>
        /// <value>
        /// The CMDS.
        /// </value>
        public IExcelCommand[] Cmds { get; private set; }

        // action - iexcelcommand[]
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandCol(Action<object> func, params IExcelCommand[] cmds)
            : this((a, b, c) => func(c), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Action<IExcelContext, Collection<string>, object> func, params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // action - action
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandCol(Action<object> func, Action command)
            : this((a, b, c) => func(c), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Action<IExcelContext, Collection<string>, object> func, Action command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // action - action<iexcelcontext>
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandCol(Action<object> func, Action<IExcelContext> command)
            : this((a, b, c) => func(c), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Action<IExcelContext, Collection<string>, object> func, Action<IExcelContext> command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<bool> - iexcelcommand[]
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandCol(Func<object, bool> func, params IExcelCommand[] cmds)
            : this((a, b, c) => func(c), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Func<IExcelContext, Collection<string>, object, bool> func, params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<bool> - action
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandCol(Func<object, bool> func, Action command)
            : this((a, b, c) => func(c), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Func<IExcelContext, Collection<string>, object, bool> func, Action command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<bool> - action<iexcelcontext>
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandCol(Func<object, bool> func, Action<IExcelContext> command)
            : this((a, b, c) => func(c), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Func<IExcelContext, Collection<string>, object, bool> func, Action<IExcelContext> command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<commandrtn> - iexcelcommand[]
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        public CommandCol(Func<object, CommandRtn> func, params IExcelCommand[] cmds)
            : this((a, b, c) => func(c), cmds) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Func<IExcelContext, Collection<string>, object, CommandRtn> func, params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<commandrtn> - action
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandCol(Func<object, CommandRtn> func, Action command)
            : this((a, b, c) => func(c), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Func<IExcelContext, Collection<string>, object, CommandRtn> func, Action command)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<commandrtn> - action<iexcelcontext>
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        public CommandCol(Func<object, CommandRtn> func, Action<IExcelContext> command)
            : this((a, b, c) => func(c), command) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandCol"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="command">The command.</param>
        /// <exception cref="ArgumentNullException">func</exception>
        public CommandCol(Func<IExcelContext, Collection<string>, object, CommandRtn> func, Action<IExcelContext> command)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Func = ExcelSerDes.DecodeFunc<IExcelContext, Collection<string>, object, CommandRtn>(r);
            Cmds = ExcelSerDes.DecodeCommands(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelSerDes.EncodeFunc(w, Func);
            ExcelSerDes.EncodeCommands(w, Cmds);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CmdCols.Push(this);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CommandCol{(When == When.Normal ? null : $"[{When}]")}: [func]"); ExcelSerDes.DescribeCommands(w, pad, Cmds); }

        static Func<IExcelContext, Collection<string>, object, CommandRtn> FuncWrapper(Action<IExcelContext, Collection<string>, object> action) => (z, c, x) => { action(z, c, x); return CommandRtn.Normal; };

        static Func<IExcelContext, Collection<string>, object, CommandRtn> FuncWrapper(Func<IExcelContext, Collection<string>, object, bool> action) => (z, c, x) => action(z, c, x) ? CommandRtn.Normal : CommandRtn.SkipCmds;

        internal static void Flush(IExcelContext ctx, int idx)
        {
            while (ctx.CmdCols.Count > idx)
                ctx.CmdCols.Pop();
        }
    }
}