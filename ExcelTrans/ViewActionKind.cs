namespace ExcelTrans
{
    /// <summary>
    /// Values for the ViewAction command
    /// </summary>
    public enum ViewActionKind
    {
        /// <summary>
        /// Freeze the columns/rows to left and above the cell
        /// </summary>
        FreezePane = 0,
        /// <summary>
        /// Sets whether the worksheet is selected within the workbook.
        /// </summary>
        SetTabSelected,
        /// <summary>
        /// Unlock all rows and columns to scroll freely
        /// </summary>
        UnfreezePane,
    }
}
