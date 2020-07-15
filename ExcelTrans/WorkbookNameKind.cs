namespace ExcelTrans
{
    /// <summary>
    /// Values for the WorkbookName command
    /// </summary>
    public enum WorkbookNameKind
    {
        /// <summary>
        /// Add a new named range
        /// </summary>
        Add = 0,
        /// <summary>
        /// Sets whether the worksheet is selected within the workbook.
        /// </summary>
        AddFormula,
        /// <summary>
        /// Add a defined name referencing value
        /// </summary>
        AddValue,
        /// <summary>
        /// Remove a defined name from the collection
        /// </summary>
        Remove,
    }
}
