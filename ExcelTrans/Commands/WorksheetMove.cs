using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Moves a Worksheet with `.Name` to the Worksheet with `.TargetName` in the current Workbook
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct WorksheetMove : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; private set; }
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; private set; }
        /// <summary>
        /// Gets the name of the target.
        /// </summary>
        /// <value>
        /// The name of the target.
        /// </value>
        public string TargetName { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetMove"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="tagetName">Name of the taget.</param>
        /// <exception cref="ArgumentNullException">
        /// name
        /// or
        /// tagetName
        /// </exception>
        public WorksheetMove(string name, string tagetName)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            TargetName = tagetName ?? throw new ArgumentNullException(nameof(tagetName));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
            TargetName = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
            w.Write(TargetName);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorksheetMove(Name, TargetName);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorksheetMove: {Name}->{TargetName}"); }
    }
}