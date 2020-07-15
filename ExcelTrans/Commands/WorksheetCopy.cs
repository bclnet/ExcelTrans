using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Copies a Worksheet with `.Name` to a new Worksheet with `.NewName` in the current Workbook
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct WorksheetCopy : IExcelCommand
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
        /// Creates new name.
        /// </summary>
        /// <value>
        /// The new name.
        /// </value>
        public string NewName { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorksheetCopy"/> struct.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="newName">The new name.</param>
        /// <exception cref="ArgumentNullException">
        /// name
        /// or
        /// newName
        /// </exception>
        public WorksheetCopy(string name, string newName)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            NewName = newName ?? throw new ArgumentNullException(nameof(newName));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
            NewName = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
            w.Write(NewName);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorksheetCopy(Name, NewName);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorksheetCopy: {Name}->{NewName}"); }
    }
}