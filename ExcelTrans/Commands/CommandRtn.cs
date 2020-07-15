using System;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Values for the return value of Commands
    /// </summary>
    [Flags]
    public enum CommandRtn
    {
        /// <summary>
        /// Normal operations.
        /// </summary>
        Normal = 0,
        /// <summary>
        /// Value should be considered a formula.
        /// </summary>
        Formula = 1,
        /// <summary>
        /// Continue to the next row.
        /// </summary>
        Continue = 2,
        /// <summary>
        /// Skip processing the attached commands.
        /// </summary>
        SkipCmds = 4,
    }
}