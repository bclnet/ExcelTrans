using ExcelTrans.Utils;
using System;
using System.Collections.ObjectModel;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Executes `.Func()` per Value
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public class CommandValue : IExcelCommand
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
        public Func<IExcelContext, Collection<string>, object, bool> Func { get; private set; }
        /// <summary>
        /// Gets the kind of the value.
        /// </summary>
        /// <value>
        /// The kind of the value.
        /// </value>
        public CellValueKind ValueKind { get; private set; }
        /// <summary>
        /// Gets the value format.
        /// </summary>
        /// <value>
        /// The value format.
        /// </value>
        public Func<IExcelContext, string> ValueFormat { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CommandValue" /> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valuePattern">The value pattern.</param>
        public CommandValue(Func<bool> func, CellValueKind valueKind, string valuePattern = null)
            : this((a, b, c) => func(), valueKind, z => valuePattern) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandValue"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CommandValue(Func<bool> func, CellValueKind valueKind, Func<string> valueFormat = null)
            : this((a, b, c) => func(), valueKind, z => valueFormat()) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandValue"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CommandValue(Func<bool> func, CellValueKind valueKind, Func<IExcelContext, string> valueFormat = null)
            : this((a, b, c) => func(), valueKind, valueFormat) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandValue"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CommandValue(Func<IExcelContext, Collection<string>, object, bool> func, CellValueKind valueKind, string valueFormat = null)
            : this(func, valueKind, z => valueFormat) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandValue"/> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CommandValue(Func<IExcelContext, Collection<string>, object, bool> func, CellValueKind valueKind, Func<string> valueFormat = null)
            : this(func, valueKind, z => valueFormat()) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CommandValue" /> class.
        /// </summary>
        /// <param name="func">The function.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <param name="valueFormat">The value format.</param>
        public CommandValue(Func<IExcelContext, Collection<string>, object, bool> func, CellValueKind valueKind, Func<IExcelContext, string> valueFormat = null)
        {
            When = When.Normal;
            Func = func;
            ValueKind = valueKind;
            ValueFormat = valueFormat;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Func = ExcelSerDes.DecodeFunc<IExcelContext, Collection<string>, object, bool>(r);
            ValueKind = (CellValueKind)r.ReadInt32();
            ValueFormat = ExcelSerDes.DecodeFunc<IExcelContext, string>(r);
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            ExcelSerDes.EncodeFunc(w, Func);
            w.Write((int)ValueKind);
            ExcelSerDes.EncodeFunc(w, ValueFormat);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CmdValues.Push(this);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CommandValue: [func]"); }

        internal static void Flush(IExcelContext ctx, int idx)
        {
            while (ctx.CmdValues.Count > idx)
                ctx.CmdValues.Pop();
        }
    }
}