namespace ExcelTrans
{
    /// <summary>
    /// Values for the WorkbookProtection command
    /// </summary>
    public enum WorkbookProtectionKind
    {
        /// <summary>
        /// Locks the structure, which prevents users from adding or deleting worksheets or from displaying hidden worksheets.
        /// </summary>
        LockStructure = 0,
        /// <summary>
        /// Locks the position of the workbook window.
        /// </summary>
        LockWindows,
        /// <summary>
        /// Lock the workbook for revision
        /// </summary>
        LockRevision,
        /// <summary>
        /// Sets a password for the workbook. This does not encrypt the workbook.
        /// </summary>
        SetPassword,
    }
}
