﻿using System.Runtime.CompilerServices;
using ExcelTrans.Commands;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml.VBA;

[assembly: InternalsVisibleTo("ExcelTrans.Tests")]

namespace ExcelTrans
{
    public interface IExcelContext : IDisposable
    {
        int XStart { get; set; }
        int X { get; set; }
        int DeltaX { get; set; }
        int Y { get; set; }
        int DeltaY { get; set; }
        int CsvX { get; set; }
        int CsvY { get; set; }
        NextDirection NextDirection { get; set; }
        Stack<CommandRow> CmdRows { get; }
        Stack<CommandCol> CmdCols { get; }
        Stack<IExcelSet> Sets { get; }
        Stack<object> Frames { get; }
        object Frame { get; set; }
        void Flush();
        ExcelRangeBase Get(string cells);
        ExcelRangeBase Next(ExcelRangeBase range, NextDirection? nextDirection = null);
        ExcelColumn Next(ExcelColumn column);
        ExcelRow Next(ExcelRow row);
    }

    internal class ExcelContext : IExcelContext
    {
        public ExcelContext()
        {
            P = new ExcelPackage();
            WB = P.Workbook;
        }
        public void Dispose() => P.Dispose();

        public int XStart { get; set; } = 1;
        public int X { get; set; } = 1;
        public int DeltaX { get; set; } = 1;
        public int Y { get; set; } = 1;
        public int DeltaY { get; set; } = 1;
        public int CsvX { get; set; } = 1;
        public int CsvY { get; set; } = 1;
        public NextDirection NextDirection { get; set; } = NextDirection.Column;
        public Stack<CommandRow> CmdRows { get; } = new Stack<CommandRow>();
        public Stack<CommandCol> CmdCols { get; } = new Stack<CommandCol>();
        public Stack<IExcelSet> Sets { get; } = new Stack<IExcelSet>();
        public Stack<object> Frames { get; } = new Stack<object>();
        public ExcelPackage P;
        public ExcelWorkbook WB;
        public ExcelWorksheet WS;
        public ExcelVbaProject V;

        public void OpenWorkbook(FileInfo path, string password = null)
        {
            P = password == null ? new ExcelPackage(path) : new ExcelPackage(path, password);
            WB = P.Workbook;
        }

        public ExcelVbaProject EnsureVba() { if (V != null) return V; WB.CreateVBAProject(); V = WB.VbaProject; return V; }

        public ExcelWorksheet EnsureWorksheet() => WS ?? (WS = WB.Worksheets.Add($"Sheet {WB.Worksheets.Count + 1}"));

        public ExcelRangeBase Get(string cells) => WS.Cells[this.DecodeAddress(cells)];

        public ExcelRangeBase Next(ExcelRangeBase range, NextDirection? nextDirection = null) => (nextDirection ?? NextDirection) == NextDirection.Column ? range.Offset(0, DeltaX) : range.Offset(DeltaY, 0);
        public ExcelColumn Next(ExcelColumn col) => null;
        public ExcelRow Next(ExcelRow row) => null;

        public void Flush()
        {
            if (Sets.Count == 0) this.WriteRowLast(null);
            Frames.Clear();
            CommandRow.Flush(this, 0);
            CommandCol.Flush(this, 0);
            PopSet.Flush(this, 0);
        }

        public object Frame
        {
            get => (CmdRows.Count, CmdCols.Count, Sets.Count);
            set
            {
                var (rows, cols, sets) = ((int rows, int cols, int sets))value;
                CommandRow.Flush(this, rows);
                CommandCol.Flush(this, cols);
                PopSet.Flush(this, sets);
            }
        }
    }
}