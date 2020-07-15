namespace ExcelTrans
{
    /// <summary>
    /// Values for the ColumnValue command
    /// </summary>
    public enum ColumnValueKind
    {
        /// <summary>
        /// Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property. Note: Cells containing formulas are ignored since EPPlus don't have a calculation engine. Wrapped and merged cells are also ignored. (set-only)
        /// </summary>
        AutoFit = 0,
        /// <summary>
        /// If set to true a column automaticlly resize(grow wider) when a user inputs numbers in a cell.
        /// </summary>
        BestFit,
        /// <summary>
        /// none
        /// </summary>
        Merged,
        /// <summary>
        /// Sets the width of the column in the worksheet
        /// </summary>
        Width,
        /// <summary>
        /// Set width to a scaled-value that should result in the nearest possible value to the true desired setting. (set-only)
        /// </summary>
        TrueWidth,
    }
}
