namespace ExcelTrans
{
    /// <summary>
    /// Values for the Protection command
    /// </summary>
    public enum ProtectionKind
    {
        /// <summary>
        /// Allow users to Format rows
        /// </summary>
        AllowFormatRows = 0,
        /// <summary>
        /// Allow users to sort a range
        /// </summary>
        AllowSort,
        /// <summary>
        /// Allow users to delete rows
        /// </summary>
        AllowDeleteRows,
        /// <summary>
        /// Allow users to delete columns
        /// </summary>
        AllowDeleteColumns,
        /// <summary>
        /// Allow users to insert hyperlinks
        /// </summary>
        AllowInsertHyperlinks,
        /// <summary>
        /// Allow users to insert rows
        /// </summary>
        AllowInsertRows,
        /// <summary>
        /// Allow users to insert columns
        /// </summary>
        AllowInsertColumns,
        /// <summary>
        /// Allow users to use autofilters
        /// </summary>
        AllowAutoFilter,
        /// <summary>
        /// Allow users to use pivottables
        /// </summary>
        AllowPivotTables,
        /// <summary>
        /// Allow users to format cells
        /// </summary>
        AllowFormatCells,
        /// <summary>
        /// Allow users to edit senarios
        /// </summary>
        AllowEditScenarios,
        /// <summary>
        /// Allow users to edit objects
        /// </summary>
        AllowEditObject,
        /// <summary>
        /// Allow users to select unlocked cells
        /// </summary>
        AllowSelectUnlockedCells,
        /// <summary>
        /// Allow users to select locked cells
        /// </summary>
        AllowSelectLockedCells,
        /// <summary>
        /// If the worksheet is protected.
        /// </summary>
        IsProtected,
        /// <summary>
        /// Allow users to Format columns
        /// </summary>
        AllowFormatColumns,
        /// <summary>
        /// Sets a password for the sheet.
        /// </summary>
        SetPassword,
    }
}
