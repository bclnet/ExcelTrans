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
        /// Continue to the next row.
        /// </summary>
        Continue = 1,
        /// <summary>
        /// Skip processing the attached commands.
        /// </summary>
        SkipCmds = 2,
    }
}