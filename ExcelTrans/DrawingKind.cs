namespace ExcelTrans
{
    /// <summary>
    /// Values for the Drawing command
    /// </summary>
    public enum DrawingKind
    {
        /// <summary>
        /// Add a new chart to the worksheet. Does not support Bubble-, Radar-, Stock- or Surface charts.
        /// </summary>
        AddChart = 0,
        /// <summary>
        /// Add a picure to the worksheet
        /// </summary>
        AddPicture,
        /// <summary>
        /// Add a new shape to the worksheet
        /// </summary>
        AddShape,
        /// <summary>
        /// Removes all drawings from the collection
        /// </summary>
        Clear,
        /// <summary>
        /// Removes a drawing.
        /// </summary>
        Remove,
    }
}
