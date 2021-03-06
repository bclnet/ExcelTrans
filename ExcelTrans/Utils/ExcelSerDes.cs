﻿using ExcelTrans.Commands;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Text;

namespace ExcelTrans.Utils
{
    internal class ExcelSerDes
    {
        public static readonly List<object> funcs = new List<object>();
        public static readonly List<Type> cmds = new List<Type>() {
            typeof(CellStyle), typeof(CellValidation), typeof(CellValue),
            typeof(ColumnValue),
            typeof(Command), typeof(CommandCol), typeof(CommandRow), typeof(CommandValue),
            typeof(ConditionalFormatting),
            typeof(Drawing),
            typeof(Flush),
            typeof(PopFrame), typeof(PopSet), typeof(PushFrame), typeof(PushSet<>),
            typeof(Protection),
            typeof(RowValue),
            typeof(VbaModule), typeof(VbaReference),
            typeof(ViewAction),
            typeof(WorkbookName), typeof(WorkbookOpen), typeof(WorkbookProtection),
            typeof(WorksheetAdd), typeof(WorksheetCopy), typeof(WorksheetDelete), typeof(WorksheetGet), typeof(WorksheetMove) };

        public static string Encode(params IExcelCommand[] cmds)
        {
            using (var b = new MemoryStream())
            using (var w = new BinaryWriter(b))
            {
                EncodeCommands(w, cmds);
                b.Position = 0;
                return Convert.ToBase64String(b.ToArray());
            }
        }

        public static string Describe(string prefix, params IExcelCommand[] cmds)
        {
            var b = new StringBuilder();
            using (var w = new StringWriter(b))
            {
                DescribeCommands(w, 0, cmds);
                return b.ToString();
            }
        }

        public static IExcelCommand[] Decode(string packed)
        {
            using (var b = new MemoryStream(Convert.FromBase64String(packed)))
            using (var r = new BinaryReader(b))
                return DecodeCommands(r);
        }

        public static void EncodeAction(BinaryWriter w, object action) { w.Write((ushort)funcs.Count); funcs.Add(action); }
        public static Action DecodeAction(BinaryReader r) => (Action)funcs[r.ReadUInt16()];
        public static Action<T1> DecodeAction<T1>(BinaryReader r) => (Action<T1>)funcs[r.ReadUInt16()];
        public static Action<T1, T2> DecodeAction<T1, T2>(BinaryReader r) => (Action<T1, T2>)funcs[r.ReadUInt16()];
        public static Action<T1, T2, T3> DecodeAction<T1, T2, T3>(BinaryReader r) => (Action<T1, T2, T3>)funcs[r.ReadUInt16()];
        public static Action<T1, T2, T3, T4> DecodeAction<T1, T2, T3, T4>(BinaryReader r) => (Action<T1, T2, T3, T4>)funcs[r.ReadUInt16()];
        public static Action<T1, T2, T3, T4, T5> DecodeAction<T1, T2, T3, T4, T5>(BinaryReader r) => (Action<T1, T2, T3, T4, T5>)funcs[r.ReadUInt16()];

        public static void EncodeFunc(BinaryWriter w, object func) { w.Write((ushort)funcs.Count); funcs.Add(func); }
        public static Func<TR> DecodeFunc<TR>(BinaryReader r) => (Func<TR>)funcs[r.ReadUInt16()];
        public static Func<T1, TR> DecodeFunc<T1, TR>(BinaryReader r) => (Func<T1, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, TR> DecodeFunc<T1, T2, TR>(BinaryReader r) => (Func<T1, T2, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, T3, TR> DecodeFunc<T1, T2, T3, TR>(BinaryReader r) => (Func<T1, T2, T3, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, T3, T4, TR> DecodeFunc<T1, T2, T3, T4, TR>(BinaryReader r) => (Func<T1, T2, T3, T4, TR>)funcs[r.ReadUInt16()];
        public static Func<T1, T2, T3, T4, T5, TR> DecodeFunc<T1, T2, T3, T4, T5, TR>(BinaryReader r) => (Func<T1, T2, T3, T4, T5, TR>)funcs[r.ReadUInt16()];

        public static void EncodeCommands(BinaryWriter w, IExcelCommand[] cmds)
        {
            if (cmds == null)
            {
                w.Write((ushort)0);
                return;
            }
            w.Write((ushort)cmds.Length);
            foreach (var cmd in cmds)
            {
                if (cmd == null) continue;
                var cmdType = cmd.GetType();
                var cmdId = ExcelSerDes.cmds.IndexOf(cmdType.GenericTypeArguments.Length == 0 ? cmdType : cmdType.GetGenericTypeDefinition());
                if (cmdId == -1) throw new InvalidOperationException($"{cmd} does not exist");
                w.Write((ushort)cmdId);
                if (cmdType.GenericTypeArguments.Length > 0)
                {
                    w.Write((ushort)cmdType.GenericTypeArguments.Length);
                    foreach (var x in cmdType.GenericTypeArguments)
                        w.Write(x.FullName);
                }
                cmd.Write(w);
            }
        }

        public static void DescribeCommands(StringWriter w, int pad, IExcelCommand[] cmds)
        {
            if (cmds == null)
                return;
            pad += 3;
            foreach (var cmd in cmds)
            {
                w.Write(ExcelService.Comment); cmd.Describe(w, pad);
            }
        }

        public static IExcelCommand[] DecodeCommands(BinaryReader r)
        {
            var cmds = new IExcelCommand[r.ReadUInt16()];
            for (var i = 0; i < cmds.Length; i++)
            {
                var cmdType = ExcelSerDes.cmds[r.ReadUInt16()];
                if (cmdType.IsGenericTypeDefinition)
                {
                    var types = new Type[r.ReadUInt16()];
                    for (var j = 0; j < types.Length; j++)
                        types[j] = Type.GetType(r.ReadString());
                    cmdType = cmdType.MakeGenericType(types);
                }
                var cmd = (IExcelCommand)FormatterServices.GetUninitializedObject(cmdType);
                cmd.Read(r);
                cmds[i] = cmd;
            }
            return cmds;
        }
    }
}