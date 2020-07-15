using System;

namespace ExcelTrans
{
    /// <summary>
    /// Values for the Address fields
    /// </summary>
    [Flags]
    public enum Address : ushort
    {
        /// <summary>
        /// Cell relative address
        /// </summary>
        Cell = CellAbs | InternalRel,
        /// <summary>
        /// Cell absolute address
        /// </summary>
        CellAbs = 1,
        /// <summary>
        /// Range relative address
        /// </summary>
        Range = RangeAbs | InternalRel,
        /// <summary>
        /// Range absoulute address
        /// </summary>
        RangeAbs = 2,
        /// <summary>
        /// Row or Column address
        /// </summary>
        RowOrCol = 3,
        /// <summary>
        /// Column to Column address
        /// </summary>
        ColToCol = 4,
        /// <summary>
        /// Row to Row addess
        /// </summary>
        RowToRow = 5,
        /// <summary>
        /// Internal
        /// </summary>
        InternalRel = 0x10,
    }
}