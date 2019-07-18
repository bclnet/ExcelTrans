using System;

namespace ExcelTrans
{
    [Flags]
    public enum Address : ushort
    {
        Cell = CellAbs | Rel,
        CellAbs = 1,
        Range = RangeAbs | Rel,
        RangeAbs = 2,
        RowOrCol = 3,
        ColToCol = 4,
        RowToRow = 5,
        Rel = 0x10, // Internal
    }
}