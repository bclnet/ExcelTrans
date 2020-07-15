using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Deletes a Worksheet with `.Name` from the current Workbook
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct WorksheetDelete : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetDelete"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <exception cref="ArgumentNullException">name</exception>
        public WorksheetDelete(string name)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        void IExcelCommand.Read(BinaryReader r) => Name = r.ReadString();

        void IExcelCommand.Write(BinaryWriter w) => w.Write(Name);

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorksheetDelete(Name);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorksheetDelete: {Name}"); }
    }
}