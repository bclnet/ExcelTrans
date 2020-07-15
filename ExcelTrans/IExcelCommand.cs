using ExcelTrans.Commands;
using System;
using System.IO;

namespace ExcelTrans
{
    /// <summary>
    /// IExcelCommand
    /// </summary>
    public interface IExcelCommand
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        When When { get; }
        /// <summary>
        /// Reads the specified r.
        /// </summary>
        /// <param name="r">The r.</param>
        void Read(BinaryReader r);
        /// <summary>
        /// Writes the specified w.
        /// </summary>
        /// <param name="w">The w.</param>
        void Write(BinaryWriter w);
        /// <summary>
        /// Executes the specified CTX.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="after">The after.</param>
        void Execute(IExcelContext ctx, ref Action after);
        /// <summary>
        /// Describes the specified w.
        /// </summary>
        /// <param name="w">The w.</param>
        /// <param name="pad">The pad.</param>
        void Describe(StringWriter w, int pad);
    }
}
