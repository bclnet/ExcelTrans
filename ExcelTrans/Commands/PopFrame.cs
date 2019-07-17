﻿using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct PopFrame : IExcelCommand
    {
        public When When { get; }
        void IExcelCommand.Read(BinaryReader r) { }
        void IExcelCommand.Write(BinaryWriter w) { }
        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Frame = ctx.Frames.Pop();

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}PopFrame"); }
    }
}