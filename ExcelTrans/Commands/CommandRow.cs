using ExcelTrans.Utils;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    public class CommandRow : IExcelCommand
    {
        public When When { get; private set; }
        public Func<IExcelContext, Collection<string>, CommandRtn> Func { get; private set; }
        public IExcelCommand[] Cmds { get; private set; }

        static Func<IExcelContext, Collection<string>, CommandRtn> FuncWrapper(Action<IExcelContext, Collection<string>> action) => (z, c) => { action(z, c); return CommandRtn.Normal; };
        static Func<IExcelContext, Collection<string>, CommandRtn> FuncWrapper(Func<IExcelContext, Collection<string>, bool> action) => (z, c) => action(z, c) ? CommandRtn.Normal : CommandRtn.SkipCmds;

        // action - iexcelcommand[]
        public CommandRow(Action func, params IExcelCommand[] cmds)
            : this(When.Before, (a, b) => func(), cmds) { }
        public CommandRow(Action<IExcelContext, Collection<string>> func, params IExcelCommand[] cmds)
            : this(When.Before, func, cmds) { }
        public CommandRow(When when, Action func, params IExcelCommand[] cmds)
            : this(when, (a, b) => func(), cmds) { }
        public CommandRow(When when, Action<IExcelContext, Collection<string>> func, params IExcelCommand[] cmds)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // action - action
        public CommandRow(Action func, Action command)
            : this(When.Before, (a, b) => func(), command) { }
        public CommandRow(Action<IExcelContext, Collection<string>> func, Action command)
            : this(When.Before, func, command) { }
        public CommandRow(When when, Action func, Action command)
            : this(when, (a, b) => func(), command) { }
        public CommandRow(When when, Action<IExcelContext, Collection<string>> func, Action command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // action - action<iexcelcontext>
        public CommandRow(Action func, Action<IExcelContext> command)
            : this(When.Before, (a, b) => func(), command) { }
        public CommandRow(Action<IExcelContext, Collection<string>> func, Action<IExcelContext> command)
            : this(When.Before, func, command) { }
        public CommandRow(When when, Action func, Action<IExcelContext> command)
            : this(when, (a, b) => func(), command) { }
        public CommandRow(When when, Action<IExcelContext, Collection<string>> func, Action<IExcelContext> command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<bool> - iexcelcommand[]
        public CommandRow(Func<bool> func, params IExcelCommand[] cmds)
            : this(When.Before, (a, b) => func(), cmds) { }
        public CommandRow(Func<IExcelContext, Collection<string>, bool> func, params IExcelCommand[] cmds)
            : this(When.Before, func, cmds) { }
        public CommandRow(When when, Func<bool> func, params IExcelCommand[] cmds)
            : this(when, (a, b) => func(), cmds) { }
        public CommandRow(When when, Func<IExcelContext, Collection<string>, bool> func, params IExcelCommand[] cmds)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<bool> - action
        public CommandRow(Func<bool> func, Action command)
            : this(When.Before, (a, b) => func(), command) { }
        public CommandRow(Func<IExcelContext, Collection<string>, bool> func, Action command)
            : this(When.Before, func, command) { }
        public CommandRow(When when, Func<bool> func, Action command)
            : this(when, (a, b) => func(), command) { }
        public CommandRow(When when, Func<IExcelContext, Collection<string>, bool> func, Action command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<bool> - action<iexcelcontext>
        public CommandRow(Func<bool> func, Action<IExcelContext> command)
            : this(When.Before, (a, b) => func(), command) { }
        public CommandRow(Func<IExcelContext, Collection<string>, bool> func, Action<IExcelContext> command)
            : this(When.Before, func, command) { }
        public CommandRow(When when, Func<bool> func, Action<IExcelContext> command)
            : this(when, (a, b) => func(), command) { }
        public CommandRow(When when, Func<IExcelContext, Collection<string>, bool> func, Action<IExcelContext> command)
        {
            When = when;
            Func = func != null ? FuncWrapper(func) : throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }

        // func<commandrtn> - iexcelcommand[]
        public CommandRow(Func<CommandRtn> func, params IExcelCommand[] cmds)
            : this(When.Before, (a, b) => func(), cmds) { }
        public CommandRow(Func<IExcelContext, Collection<string>, CommandRtn> func, params IExcelCommand[] cmds)
            : this(When.Before, func, cmds) { }
        public CommandRow(When when, Func<CommandRtn> func, params IExcelCommand[] cmds)
            : this(when, (a, b) => func(), cmds) { }
        public CommandRow(When when, Func<IExcelContext, Collection<string>, CommandRtn> func, params IExcelCommand[] cmds)
        {
            When = when;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = cmds;
        }
        // func<commandrtn> - action
        public CommandRow(Func<CommandRtn> func, Action command)
            : this(When.Before, (a, b) => func(), command) { }
        public CommandRow(Func<IExcelContext, Collection<string>, CommandRtn> func, Action command)
            : this(When.Before, func, command) { }
        public CommandRow(When when, Func<CommandRtn> func, Action command)
            : this(when, (a, b) => func(), command) { }
        public CommandRow(When when, Func<IExcelContext, Collection<string>, CommandRtn> func, Action command)
        {
            When = when;
            Func = func ?? throw new ArgumentNullException(nameof(func));
            Cmds = new[] { new Command(command) };
        }
        // func<commandrtn> - action<iexcelcontext>
        public CommandRow(Func<CommandRtn> func, Action<IExcelContext> command)
            : this(When.Before, (a, b) => func(), command) { }
        public CommandRow(Func<IExcelContext, Collection<string>, CommandRtn> func, Action<IExcelContext> command)
            : this(When.Before, func, command) { }
        public CommandRow(When when, Func<CommandRtn> func, Action<IExcelContext> command)
            : this(when, (a, b) => func(), command) { }
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

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CommandRow{(When == When.Before ? null : $"[{When}]")}: [func]"); ExcelSerDes.DescribeCommands(w, pad, Cmds); }

        internal static void Flush(IExcelContext ctx, int idx)
        {
            while (ctx.CmdRows.Count > idx)
                ctx.CmdRows.Pop();
        }
    }
}