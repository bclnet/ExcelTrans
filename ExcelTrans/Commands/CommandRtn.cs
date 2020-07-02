using System;

namespace ExcelTrans.Commands
{
    [Flags]
    public enum CommandRtn
    {
        Normal = 0,
        Formula = 1,
        Continue = 2,
        SkipCmds = 4,
    }
}