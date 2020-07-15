using System;
using System.IO;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Opens a Workbook at `.Path` with optional `.Password`
    /// </summary>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    public struct WorkbookOpen : IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the path.
        /// </summary>
        /// <value>
        /// The path.
        /// </value>
        public string Path { get; private set; }
        /// <summary>
        /// Gets the password.
        /// </summary>
        /// <value>
        /// The password.
        /// </value>
        public string Password { get; private set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="WorkbookOpen"/> struct.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="password">The password.</param>
        /// <exception cref="ArgumentNullException">path</exception>
        public WorkbookOpen(string path, string password = null)
        {
            When = When.Normal;
            Path = path ?? throw new ArgumentNullException(nameof(path));
            Password = password;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Path = r.ReadString();
            Password = r.ReadBoolean() ? r.ReadString() : null;
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Path);
            w.Write(Password != null); if (Password != null) w.Write(Password);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorkbookOpen(Path, Password);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorkbookOpen: {Path}"); }
    }
}