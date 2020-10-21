using ExcelTrans.Commands;
using OfficeOpenXml;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
[assembly: InternalsVisibleTo("ExcelTrans.Tests")]

namespace ExcelTrans
{
    /// <summary>
    /// IExcelContext
    /// </summary>
    /// <seealso cref="System.IDisposable" />
    public interface IExcelContext : IDisposable
    {
        /// <summary>
        /// Gets a value indicating whether macros are disabled.
        /// </summary>
        /// <value>
        ///   <c>true</c> if macros are disabled otherwise, <c>false</c>.
        /// </value>
        bool MacrosDisabled { get; }
        /// <summary>
        /// Gets or sets the <see cref="System.Object"/> with the specified row.
        /// </summary>
        /// <value>
        /// The <see cref="System.Object"/>.
        /// </value>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <returns></returns>
        object this[int row, int col] { get; set; }
        /// <summary>
        /// Gets or sets the <see cref="System.Object"/> with the specified cell.
        /// </summary>
        /// <value>
        /// The <see cref="System.Object"/>.
        /// </value>
        /// <param name="cell">The cell.</param>
        /// <returns></returns>
        object this[(int row, int col) cell] { get; set; }
        /// <summary>
        /// Gets or sets where the cursor X starts per row.
        /// </summary>
        /// <value>
        /// The x start.
        /// </value>
        int XStart { get; set; }
        /// <summary>
        /// Gets or sets the cursor X coordinate.
        /// </summary>
        /// <value>
        /// The x.
        /// </value>
        int X { get; set; }
        /// <summary>
        /// Gets or sets the cursor Y coordinate.
        /// </summary>
        /// <value>
        /// The y.
        /// </value>
        int Y { get; set; }
        /// <summary>
        /// Gets or sets the amount the cursor X advances.
        /// </summary>
        /// <value>
        /// The delta x.
        /// </value>
        int DeltaX { get; set; }
        /// <summary>
        /// Gets or sets the amount the cursor Y advances.
        /// </summary>
        /// <value>
        /// The delta y.
        /// </value>
        int DeltaY { get; set; }
        /// <summary>
        /// Gets or sets the cursor CsvX coordinate, advances with X.
        /// </summary>
        /// <value>
        /// The CSV x.
        /// </value>
        int CsvX { get; set; }
        /// <summary>
        /// Gets or sets the cursor CsvY coordinate, advances with Y.
        /// </summary>
        /// <value>
        /// The CSV y.
        /// </value>
        int CsvY { get; set; }
        /// <summary>
        /// Gets or sets the next direction.
        /// </summary>
        /// <value>
        /// The next direction.
        /// </value>
        NextDirection NextDirection { get; set; }
        /// <summary>
        /// Gets the stack of commands per row.
        /// </summary>
        /// <value>
        /// The command rows.
        /// </value>
        Stack<CommandRow> CmdRows { get; }
        /// <summary>
        /// Gets the stack of commands per column.
        /// </summary>
        /// <value>
        /// The command cols.
        /// </value>
        Stack<CommandCol> CmdCols { get; }
        /// <summary>
        /// Gets the stack of commands per value.
        /// </summary>
        /// <value>
        /// The command values.
        /// </value>
        Stack<CommandValue> CmdValues { get; }
        /// <summary>
        /// Gets the stack of sets.
        /// </summary>
        /// <value>
        /// The sets.
        /// </value>
        Stack<IExcelSet> Sets { get; }
        /// <summary>
        /// Gets the stack of frames.
        /// </summary>
        /// <value>
        /// The frames.
        /// </value>
        Stack<object> Frames { get; }
        /// <summary>
        /// Gets the current frame.
        /// </summary>
        /// <value>
        /// The frame.
        /// </value>
        object Frame { get; set; }
        /// <summary>
        /// Flushes all pending commands.
        /// </summary>
        void Flush();
        /// <summary>
        /// Gets the specified range.
        /// </summary>
        /// <param name="cells">The cells.</param>
        /// <returns></returns>
        ExcelRangeBase Get(string cells);
        /// <summary>
        /// Gets the specified row.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns></returns>
        T Get<T>(int row, int col, T defaultValue = default);
        /// <summary>
        /// Gets the specified row.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="cell">The cell.</param>
        /// <param name="defaultValue">The default value.</param>
        /// <returns></returns>
        T Get<T>((int row, int col) cell, T defaultValue = default);
        /// <summary>
        /// Advances the cursor based on NextDirection.
        /// </summary>
        /// <param name="range">The range.</param>
        /// <param name="nextDirection">The next direction.</param>
        /// <returns></returns>
        ExcelRangeBase Next(ExcelRangeBase range, NextDirection? nextDirection = null);
        /// <summary>
        /// Advances the cursor to the next row.
        /// </summary>
        /// <param name="column">The column.</param>
        /// <returns></returns>
        ExcelColumn Next(ExcelColumn column);
        /// <summary>
        /// Advances the cursor to the next row.
        /// </summary>
        /// <param name="row">The row.</param>
        /// <returns></returns>
        ExcelRow Next(ExcelRow row);
    }

    internal class ExcelContext : IExcelContext
    {
        public ExcelContext(bool macrosDisabled = false)
        {
            MacrosDisabled = macrosDisabled;
            P = new ExcelPackage();
            WB = P.Workbook;
        }
        public void Dispose() => P.Dispose();

        public object this[int row, int col]
        {
            get => this.GetCellValue(row, col);
            set => this.CellValue(row, col, value);
        }
        public object this[(int row, int col) cell]
        {
            get => this.GetCellValue(cell);
            set => this.CellValue(cell, value);
        }

        public bool MacrosDisabled { get; }
        public int XStart { get; set; } = 1;
        public int X { get; set; } = 1;
        public int Y { get; set; } = 1;
        public int DeltaX { get; set; } = 1;
        public int DeltaY { get; set; } = 1;
        public int CsvX { get; set; } = 1;
        public int CsvY { get; set; } = 1;
        public NextDirection NextDirection { get; set; } = NextDirection.Column;
        public Stack<CommandRow> CmdRows { get; } = new Stack<CommandRow>();
        public Stack<CommandCol> CmdCols { get; } = new Stack<CommandCol>();
        public Stack<CommandValue> CmdValues { get; } = new Stack<CommandValue>();
        public Stack<IExcelSet> Sets { get; } = new Stack<IExcelSet>();
        public Stack<object> Frames { get; } = new Stack<object>();
        public ExcelPackage P;
        public ExcelWorkbook WB;
        public ExcelWorksheet WS;
        public ExcelVbaProject V;

        public ExcelVbaProject EnsureVba()
        {
            if (V != null) return V;
            if (MacrosDisabled) throw new InvalidOperationException("Macros are disabled");
            WB.CreateVBAProject();
            V = WB.VbaProject;
            return V;
        }

        public ExcelWorksheet EnsureWorksheet() => WS ?? (WS = WB.Worksheets.Add($"Sheet {WB.Worksheets.Count + 1}"));

        public ExcelRangeBase Get(string cells) => WS.Cells[this.DecodeAddress(cells)];
        public T Get<T>(int row, int col, T defaultValue = default) => this.GetCellValue(row, col) is T value ? value : defaultValue;
        public T Get<T>((int row, int col) cell, T defaultValue = default) => this.GetCellValue(cell) is T value ? value : defaultValue;

        public ExcelRangeBase Next(ExcelRangeBase range, NextDirection? nextDirection = null) => (nextDirection ?? NextDirection) == NextDirection.Column ? range.Offset(0, DeltaX) : range.Offset(DeltaY, 0);
        public ExcelColumn Next(ExcelColumn col) => throw new NotImplementedException();
        public ExcelRow Next(ExcelRow row) => throw new NotImplementedException();

        public void Flush()
        {
            if (Sets.Count == 0) this.WriteRowLast(null);
            Frames.Clear();
            CommandRow.Flush(this, 0);
            CommandCol.Flush(this, 0);
            CommandValue.Flush(this, 0);
            PopSet.Flush(this, 0);
        }

        public object Frame
        {
            get => (CmdRows.Count, CmdCols.Count, CmdValues.Count, Sets.Count);
            set
            {
                var (rows, cols, values, sets) = ((int rows, int cols, int values, int sets))value;
                CommandRow.Flush(this, rows);
                CommandCol.Flush(this, cols);
                CommandValue.Flush(this, values);
                PopSet.Flush(this, sets);
            }
        }
    }
}