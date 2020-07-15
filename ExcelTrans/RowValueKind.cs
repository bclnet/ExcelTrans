namespace ExcelTrans
{
    /// <summary>
    /// Values for the RowValue command
    /// </summary>
    public enum RowValueKind
    {
        /// <summary>
        /// If outline level is set this tells that the row is collapsed
        /// </summary>
        Collapsed = 0,
        /// <summary>
        /// Set to true if You do not want the row to Autosize
        /// </summary>
        CustomHeight,
        /// <summary>
        /// Sets the height of the row
        /// </summary>
        Height,
        /// <summary>
        /// Allows the row to be hidden in the worksheet
        /// </summary>
        Hidden,
        /// <summary>
        /// Sets the merged row
        /// </summary>
        Merged,
        /// <summary>
        /// Outline level.
        /// </summary>
        OutlineLevel,
        /// <summary>
        /// Adds a manual page break after the row.
        /// </summary>
        PageBreak,
        /// <summary>
        /// Show phonetic Information
        /// </summary>
        Phonetic,
        /// <summary>
        /// Sets the style for the entire column using a style name.
        /// </summary>
        StyleName,
    }
}
