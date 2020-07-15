using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Applies `.Value` of `.ActionKind` to the active spreadsheet
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct ViewAction : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public string Value { get; private set; }
        /// <summary>
        /// Gets the kind of the action.
        /// </summary>
        /// <value>
        /// The kind of the action.
        /// </value>
        public ViewActionKind ActionKind { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="ViewAction"/> struct.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="actionKind">Kind of the action.</param>
        public ViewAction(int row, int col, ViewActionKind actionKind)
            : this(ExcelService.GetAddress(row, col), actionKind) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ViewAction"/> struct.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="actionKind">Kind of the action.</param>
        public ViewAction(string value, ViewActionKind actionKind)
        {
            When = When.Normal;
            Value = value;
            ActionKind = actionKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Value = r.ReadBoolean() ? r.ReadString() : null;
            ActionKind = (ViewActionKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Value != null); if (Value != null) w.Write(Value);
            w.Write((int)ActionKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.ViewAction(Value, ActionKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}ViewAction: {Value} - {ActionKind}"); }
    }
}