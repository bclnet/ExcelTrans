namespace ExcelTrans
{
    /// <summary>
    /// Values for the CellValue command
    /// </summary>
    public enum CellValueKind
    {
        /// <summary>
        /// Set the range to a specific value
        /// </summary>
        Value = 0,
        /// <summary>
        /// Returns the formatted value.
        /// </summary>
        Text,
        /// <summary>
        /// Set an autofilter for the range
        /// </summary>
        AutoFilter,
        /// <summary>
        /// Set the column width from the content of the range. The minimum width is the value of the ExcelWorksheet.defaultColumnWidth property. Note: Cells containing formulas must be calculated before autofit is called. Wrapped and merged cells are also ignored. (set-only)
        /// </summary>
        AutoFitColumns,
        /// <summary>
        /// The comment text
        /// </summary>
        Comment,
        //CommentMore,
        //ConditionalFormattingMore,
        /// <summary>
        /// Copies the range of cells to an other range
        /// </summary>
        Copy,
        /// <summary>
        /// Data validation for this range (get-only)
        /// </summary>
        DataValidation,
        /// <summary>
        /// Gets or sets a formula for a range.
        /// </summary>
        Formula,
        /// <summary>
        /// Gets or Set a formula in R1C1 format.
        /// </summary>
        FormulaR1C1,
        /// <summary>
        /// Set the hyperlink property for a range of cells
        /// </summary>
        Hyperlink,
        /// <summary>
        /// If the cells in the range are merged.
        /// </summary>
        Merge,
        /// <summary>
        /// Add a rich text string
        /// </summary>
        RichText,
        /// <summary>
        /// Clear the collection
        /// </summary>
        RichTextClear,
        /// <summary>
        /// The named style
        /// </summary>
        StyleName,
    }
}
