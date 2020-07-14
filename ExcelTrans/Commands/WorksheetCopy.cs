﻿using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct WorksheetCopy : IExcelCommand
    {
        public When When { get; private set; }
        public string Name { get; private set; }
        public string NewName { get; private set; }

        public WorksheetCopy(string name, string newName)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            NewName = newName ?? throw new ArgumentNullException(nameof(newName));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
            NewName = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
            w.Write(NewName);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorksheetCopy(Name, NewName);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorksheetCopy: {Name}->{NewName}"); }
    }
}