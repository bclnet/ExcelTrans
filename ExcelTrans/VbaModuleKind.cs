namespace ExcelTrans
{
    /// <summary>
    /// Values for the VbaCodeModule command
    /// </summary>
    public enum VbaModuleKind
    {
        /// <summary>
        /// Sets the Workbook VBA Module
        /// </summary>
        CodeModule = 0,
        /// <summary>
        /// Adds a new VBA Module
        /// </summary>
        Module,
        /// <summary>
        /// Adds a new VBA public class
        /// </summary>
        Class,
        /// <summary>
        /// Adds a new VBA private class
        /// </summary>
        PrivateClass,
    }
}
