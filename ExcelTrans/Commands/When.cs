using System;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Values for the When fields of Commands
    /// </summary>
    [Flags]
    public enum When : byte
    {
        /// <summary>
        /// Execute on the first set in PushSet.
        /// </summary>
        FirstSet = 1,
        /// <summary>
        /// Execute first before rows.
        /// </summary>
        First = 2,
        /// <summary>
        /// Execute before normal writing.
        /// </summary>
        Normal = 4,
        /// <summary>
        /// Execute after normal writing.
        /// </summary>
        AfterNormal = 8,
        /// <summary>
        /// Execute last after rows.
        /// </summary>
        Last = 16,
        /// <summary>
        /// Execute on the last set in PushSet.
        /// </summary>
        LastSet = 32,
    }
}