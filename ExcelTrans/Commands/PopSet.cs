﻿using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct PopSet : IExcelCommand
    {
        public When When { get; }
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Sets.Pop().Execute(ctx);

        internal static void Flush(IExcelContext ctx, int idx)
        {
            while (ctx.Sets.Count > idx)
                ctx.Sets.Pop().Execute(ctx);
        }

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PopSet"); }
    }
}