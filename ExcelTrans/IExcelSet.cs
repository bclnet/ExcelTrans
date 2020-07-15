using System.Collections.ObjectModel;

namespace ExcelTrans
{
    /// <summary>
    /// IExcelSet
    /// </summary>
    public interface IExcelSet
    {
        /// <summary>
        /// Adds the specified s.
        /// </summary>
        /// <param name="s">The s.</param>
        void Add(Collection<string> s);
        /// <summary>
        /// Executes the specified CTX.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        void Execute(IExcelContext ctx);
    }
}
