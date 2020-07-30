namespace ExcelTrans
{
    /// <summary>
    /// Values for the VbaCodeModule command
    /// </summary>
    public enum VbaModuleKind
    {
        /// <summary>
        /// Gets or adds the VBA Module (Name:null for the Workbook VBA Module)
        /// </summary>
        Get = 0,
        /// <summary>
        /// Gets the Workbook VBA Module
        /// </summary>
        CodeModule,
        /// <summary>
        /// Adds a new VBA Module
        /// </summary>
        AddModule,
        /// <summary>
        /// Adds a new VBA public class
        /// </summary>
        AddClass,
        /// <summary>
        /// Adds a new VBA private class
        /// </summary>
        AddPrivateClass,
    }
}
