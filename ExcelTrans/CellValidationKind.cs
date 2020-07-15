namespace ExcelTrans
{
    /// <summary>
    /// Values for the CellValidation command
    /// </summary>
    public enum CellValidationKind
    {
        /// <summary>
        /// Returns the first matching validation.
        /// </summary>
        Find = 0,
        /// <summary>
        /// Adds a IExcelDataValidationAny to the worksheet.
        /// </summary>
        AnyValidation,
        /// <summary>
        /// Adds a IExcelDataValidationCustom to the worksheet.
        /// </summary>
        CustomValidation,
        /// <summary>
        /// Adds an IExcelDataValidationDateTime to the worksheet. The only accepted values are DateTime values.
        /// </summary>
        DateTimeValidation,
        /// <summary>
        /// Adds an IExcelDataValidationDecimal to the worksheet. The only accepted values are decimal values.
        /// </summary>
        DecimalValidation,
        /// <summary>
        /// Adds an IExcelDataValidationInt to the worksheet. The only accepted values are integer values.
        /// </summary>
        IntegerValidation,
        /// <summary>
        /// Adds an IExcelDataValidationList to the worksheet. The accepted values are defined in a list.
        /// </summary>
        ListValidation,
        /// <summary>
        /// Adds an IExcelDataValidationInt regarding text length to the worksheet.
        /// </summary>
        TextLengthValidation,
        /// <summary>
        /// Adds an IExcelDataValidationTime to the worksheet. The only accepted values are Time values.
        /// </summary>
        TimeValidation,
    }
}
