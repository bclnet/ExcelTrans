using ExcelTrans.Utils;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    public class CommandCol : IExcelCommand
    {
        public When When { get; }
        public Func<IExcelContext, Collection<string>, object, CommandRtn> Func { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        static Func<IExcelContext, Collection<string>, object, CommandRtn> FuncWrapper(Action<IExcelContext, Collection<string>, object> action) => (z, c, x) => { action(z, c, x); return CommandRtn.Normal; };
        static Func<IExcelContext, Collection<string>, object, CommandRtn> FuncWrapper(Func<IExcelContext, Collection<string>, object, bool> action) => (z, c, x) => action(z, c, x) ? CommandRtn.Normal : CommandRtn.SkipCmds;

        // action - iexcelcommand[]
        public CommandCol(Action<object> func, params IExcelCommand[] cmds)
            : this((a, b, c) => func(c), cmds) { }
        public CommandCol(Action<IExcelContext, Collection<string>, object> func, params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // action - action
        public CommandCol(Action<object> func, Action command)
            : this((a, b, c) => func(c), command) { }
        public CommandCol(Action<IExcelContext, Collection<string>, object> func, Action command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // action - action<iexcelcontext>
        public CommandCol(Action<object> func, Action<IExcelContext> command)
            : this((a, b, c) => func(c), command) { }
        public CommandCol(Action<IExcelContext, Collection<string>, object> func, Action<IExcelContext> command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<bool> - iexcelcommand[]
        public CommandCol(Func<object, bool> func, params IExcelCommand[] cmds)
            : this((a, b, c) => func(c), cmds) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, bool> func, params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<bool> - action
        public CommandCol(Func<object, bool> func, Action command)
            : this((a, b, c) => func(c), command) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, bool> func, Action command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<bool> - action<iexcelcontext>
        public CommandCol(Func<object, bool> func, Action<IExcelContext> command)
            : this((a, b, c) => func(c), command) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, bool> func, Action<IExcelContext> command)
        {
            When = When.Normal;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<commandrtn> - iexcelcommand[]
        public CommandCol(Func<object, CommandRtn> func, params IExcelCommand[] cmds)
            : this((a, b, c) => func(c), cmds) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, CommandRtn> func, params IExcelCommand[] cmds)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<commandrtn> - action
        public CommandCol(Func<object, CommandRtn> func, Action command)
            : this((a, b, c) => func(c), command) { }
        public CommandCol(Func<IExcelContext, Collection<string>, object, CommandRtn> func, Action command)
        {
            When = When.Normal;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<commandrtn> - action<iexcelcontext>
        public CommandCol(Func<object, CommandRtn> func, Action<IExcelContext> command)
            : this((a, b, c) => func(c), command) { }
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

        internal static void Flush(IExcelContext ctx, int idx)
        {
            while (ctx.CmdCols.Count > idx)
                ctx.CmdCols.Pop();
        }
    }
}