using ExcelTrans.Commands;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.VBA;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Xml;

namespace ExcelTrans
{
    /// <summary>
    /// ExcelExtensions
    /// </summary>
    public static class ExcelExtensions
    {
        #region Execute

        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <param name="after">The after.</param>
        /// <returns></returns>
        public static object ExecuteCmd(this IExcelContext ctx, IExcelCommand[] cmds, out Action after)
        {
            var frame = ctx.Frame;
            var afterActions = new List<Action>();
            Action action2 = null;
            foreach (var cmd in cmds)
                if (cmd == null) { }
                else if (cmd.When <= When.Normal) { cmd.Execute(ctx, ref action2); if (action2 != null) { afterActions.Add(action2); action2 = null; } }
                else afterActions.Add(() => { cmd.Execute(ctx, ref action2); if (action2 != null) { afterActions.Add(action2); action2 = null; } });
            after = afterActions.Count > 0 ? () => { foreach (var action in afterActions) action?.Invoke(); } : (Action)null;
            return frame;
        }

        /// <summary>
        /// Executes the row.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="when">The when.</param>
        /// <param name="s">The s.</param>
        /// <param name="after">The after.</param>
        /// <returns></returns>
        public static CommandRtn ExecuteRow(this IExcelContext ctx, When when, Collection<string> s, out Action after)
        {
            var cr = CommandRtn.Normal;
            var afterActions = new List<Action>();
            foreach (var cmd in ctx.CmdRows.Where(x => (x.When & when) == when))
            {
                if (cmd == null) continue;
                var r = cmd.Func(ctx, s);
                if (cmd.Cmds != null && cmd.Cmds.Length > 0 && (r & CommandRtn.SkipCmds) != CommandRtn.SkipCmds)
                {
                    ctx.Frame = ctx.ExecuteCmd(cmd.Cmds, out var action);
                    if (action != null) afterActions.Add(action);
                }
                cr |= r;
            }
            after = afterActions.Count > 0 ? () => { foreach (var action in afterActions) action.Invoke(); } : (Action)null;
            return cr;
        }

        /// <summary>
        /// Executes the col.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="s">The s.</param>
        /// <param name="v">The v.</param>
        /// <param name="after">The after.</param>
        /// <returns></returns>
        public static CommandRtn ExecuteCol(this IExcelContext ctx, Collection<string> s, object v, out Action after)
        {
            var cr = CommandRtn.Normal;
            var afterActions = new List<Action>();
            foreach (var cmd in ctx.CmdCols)
            {
                if (cmd == null) continue;
                var r = cmd.Func(ctx, s, v);
                if (cmd.Cmds != null && cmd.Cmds.Length > 0 && (r & CommandRtn.SkipCmds) != CommandRtn.SkipCmds)
                {
                    ctx.Frame = ctx.ExecuteCmd(cmd.Cmds, out var action);
                    if (action != null) afterActions.Add(action);
                }
                cr |= r;
            }
            after = afterActions.Count > 0 ? () => { foreach (var action in afterActions) action?.Invoke(); } : (Action)null;
            return cr;
        }

        #endregion

        #region Write

        /// <summary>
        /// Writes the row first set.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="s">The s.</param>
        public static void WriteRowFirstSet(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.FirstSet, s, out var after);
        /// <summary>
        /// Writes the row first.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="s">The s.</param>
        public static void WriteRowFirst(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.First, s, out var after);

        /// <summary>
        /// Advances the row.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        public static void AdvanceRow(this IExcelContext ctx) => ctx.CsvY++;
        /// <summary>
        /// Writes the row.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="s">The s.</param>
        /// <param name="startIndex">The start index.</param>
        public static void WriteRow(this IExcelContext ctx, Collection<string> s, int startIndex = 0)
        {
            var ws = ((ExcelContext)ctx).EnsureWorksheet();
            // execute-row-before
            var cr = ctx.ExecuteRow(When.Normal, s, out var after);
            if ((cr & CommandRtn.Continue) == CommandRtn.Continue)
                return;
            //
            ctx.X = ctx.XStart;
            for (var i = startIndex; i < s.Count; i++)
            {
                ctx.CsvX = i + 1;
                var v = s[i].ParseValue();
                // execute-col
                cr = ctx.ExecuteCol(s, v, out var action);
                if ((cr & CommandRtn.Continue) == CommandRtn.Continue)
                    continue;
                if (ctx.Y > 0 && ctx.X > 0)
                {
                    if ((cr & CommandRtn.Formula) != CommandRtn.Formula) ws.SetValue(ctx.Y, ctx.X, v);
                    else ws.Cells[ctx.Y, ctx.X].Formula = s[i];
                    //if (v is DateTime) ws.Cells[ExcelCellBase.GetAddress(ctx.Y, ctx.X)].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                }
                ctx.X += ctx.DeltaX;
                action?.Invoke();
            }
            after?.Invoke();
            ctx.Y += ctx.DeltaY;
            // execute-row-after
            ctx.ExecuteRow(When.AfterNormal, s, out var after2);
        }

        /// <summary>
        /// Writes the row last.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="s">The s.</param>
        public static void WriteRowLast(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.Last, s, out var after);
        /// <summary>
        /// Writes the row last set.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="s">The s.</param>
        public static void WriteRowLastSet(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.LastSet, s, out var after);

        #endregion

        #region Vba

        /// <summary>
        /// Vbas the code module.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="code">The code.</param>
        /// <param name="moduleKind">Kind of the module.</param>
        /// <exception cref="ArgumentOutOfRangeException">moduleKind</exception>
        public static void VbaCodeModule(this IExcelContext ctx, string name, VbaCode code, VbaModuleKind moduleKind)
        {
            //if (!string.IsNullOrEmpty(name) && char.IsDigit(name[0]))
            //    name = $"{name}";
            var v = ((ExcelContext)ctx).EnsureVba();
            ExcelVBAModule m;
            switch (moduleKind)
            {
                case VbaModuleKind.CodeModule: m = ((ExcelContext)ctx).WB.CodeModule; break;
                case VbaModuleKind.Module: m = v.Modules.AddModule(name); break;
                case VbaModuleKind.Class: m = v.Modules.AddClass(name, true); break;
                case VbaModuleKind.PrivateClass: m = v.Modules.AddClass(name, false); break;
                default: throw new ArgumentOutOfRangeException(nameof(moduleKind), moduleKind.ToString());
            }
            if (!string.IsNullOrEmpty(code.Name)) m.Name = code.Name;
            if (code.Description != null) m.Description = code.Description;
            if (code.Code != null) m.Code = code.ProcessCode();
            if (code.ReadOnly != null) m.ReadOnly = code.ReadOnly.Value;
            if (code.Private != null) m.Private = code.Private.Value;
        }

        /// <summary>
        /// Vbas the module.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="code">The code.</param>
        public static void VbaModule(this IExcelContext ctx, string name, VbaCode code)
        {
            //if (!string.IsNullOrEmpty(name) && char.IsDigit(name[0]))
            //    name = $"{name}";
            var v = ((ExcelContext)ctx).EnsureVba();
            var m = v.Modules[name] ?? v.Modules.AddModule(name);
            if (!string.IsNullOrEmpty(code.Name)) m.Name = code.Name;
            if (code.Description != null) m.Description = code.Description;
            if (code.Code != null) m.Code = code.ProcessCode();
            if (code.ReadOnly != null) m.ReadOnly = code.ReadOnly.Value;
            if (code.Private != null) m.Private = code.Private.Value;
        }

        /// <summary>
        /// Vbas the reference.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="libraries">The libraries.</param>
        public static void VbaReference(this IExcelContext ctx, VbaLibrary[] libraries)
        {
            var references = ((ExcelContext)ctx).EnsureVba().References;
            foreach (var library in libraries)
                references.Add(new ExcelVbaReference { Name = library.Name, Libid = library.Libid.ToString() });
        }

        /// <summary>
        /// Vbas the signature.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        public static void VbaSignature(this IExcelContext ctx)
        {
            var v = ((ExcelContext)ctx).EnsureVba();
            v.Signature.Certificate = null;
        }

        /// <summary>
        /// Vbas the protection.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        public static void VbaProtection(this IExcelContext ctx)
        {
            var v = ((ExcelContext)ctx).EnsureVba();
            v.Protection.SetPassword("EPPlus");
        }

        #endregion

        #region Workbook

        /// <summary>
        /// Workbooks the protection.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="value">The value.</param>
        /// <param name="protectionKind">Kind of the protection.</param>
        /// <exception cref="ArgumentOutOfRangeException">protectionKind</exception>
        public static void WorkbookProtection(this IExcelContext ctx, object value, WorkbookProtectionKind protectionKind)
        {
            var protection = ((ExcelContext)ctx).WB.Protection;
            switch (protectionKind)
            {
                case WorkbookProtectionKind.LockStructure: protection.LockStructure = value.CastValue<bool>(); break;
                case WorkbookProtectionKind.LockWindows: protection.LockWindows = value.CastValue<bool>(); break;
                case WorkbookProtectionKind.LockRevision: protection.LockRevision = value.CastValue<bool>(); break;
                case WorkbookProtectionKind.SetPassword: protection.SetPassword(value.CastValue<string>()); break;
                default: throw new ArgumentOutOfRangeException(nameof(protectionKind));
            }
        }

        /// <summary>
        /// Workbooks the name.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public static void WorkbookName(this IExcelContext ctx, string name, int row, int col, WorkbookNameKind nameKind = WorkbookNameKind.Add) => WorkbookName(ctx, name, ExcelService.GetAddress(row, col), nameKind);
        /// <summary>
        /// Workbooks the name.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public static void WorkbookName(this IExcelContext ctx, string name, int fromRow, int fromCol, int toRow, int toCol, WorkbookNameKind nameKind = WorkbookNameKind.Add) => WorkbookName(ctx, name, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), nameKind);
        /// <summary>
        /// Workbooks the name.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="r">The r.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public static void WorkbookName(this IExcelContext ctx, string name, Address r, WorkbookNameKind nameKind = WorkbookNameKind.Add) => WorkbookName(ctx, name, ExcelService.GetAddress(r, 0, 0), nameKind);
        /// <summary>
        /// Workbooks the name.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public static void WorkbookName(this IExcelContext ctx, string name, Address r, int row, int col, WorkbookNameKind nameKind = WorkbookNameKind.Add) => WorkbookName(ctx, name, ExcelService.GetAddress(r, row, col), nameKind);
        /// <summary>
        /// Workbooks the name.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="nameKind">Kind of the name.</param>
        public static void WorkbookName(this IExcelContext ctx, string name, Address r, int fromRow, int fromCol, int toRow, int toCol, WorkbookNameKind nameKind = WorkbookNameKind.Add) => WorkbookName(ctx, name, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), nameKind);
        /// <summary>
        /// Workbooks the name.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="cells">The cells.</param>
        /// <param name="nameKind">Kind of the name.</param>
        /// <exception cref="ArgumentOutOfRangeException">nameKind</exception>
        public static void WorkbookName(this IExcelContext ctx, string name, string cells, WorkbookNameKind nameKind = WorkbookNameKind.Add)
        {
            var names = ((ExcelContext)ctx).WB.Names;
            switch (nameKind)
            {
                case WorkbookNameKind.Add: var range = ctx.Get(cells); names.Add(name, range); break;
                case WorkbookNameKind.AddFormula: names.AddFormula(name, cells); break;
                case WorkbookNameKind.AddValue: names.AddValue(name, cells); break;
                case WorkbookNameKind.Remove: names.Remove(name); break;
                default: throw new ArgumentOutOfRangeException(nameof(nameKind));
            }
        }

        /// <summary>
        /// Workbooks the open.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="path">The path.</param>
        /// <param name="password">The password.</param>
        public static void WorkbookOpen(this IExcelContext ctx, string path, string password = null) => WorkbookOpen(ctx, new FileInfo(path), password);
        /// <summary>
        /// Workbooks the open.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="path">The path.</param>
        /// <param name="password">The password.</param>
        public static void WorkbookOpen(this IExcelContext ctx, FileInfo path, string password = null)
        {
            var ctx2 = (ExcelContext)ctx;
            ctx2.P = password == null ? new ExcelPackage(path) : new ExcelPackage(path, password);
            ctx2.WB = ctx2.P.Workbook;
        }

        #endregion

        #region Worksheet

        /// <summary>
        /// Worksheets the add.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        public static void WorksheetAdd(this IExcelContext ctx, string name)
        {
            ctx.Flush();
            var ctx2 = (ExcelContext)ctx;
            ctx2.WS = ctx2.WB.Worksheets.Add(name);
            ctx.DeltaX = ctx.DeltaY = ctx.XStart = ctx.Y = 1;
        }

        /// <summary>
        /// Worksheets the copy.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="newName">The new name.</param>
        public static void WorksheetCopy(this IExcelContext ctx, string name, string newName)
        {
            ctx.Flush();
            var ctx2 = (ExcelContext)ctx;
            ctx2.WS = ctx2.WB.Worksheets.Copy(name, newName);
            ctx.DeltaX = ctx.DeltaY = ctx.XStart = ctx.Y = 1;
        }

        /// <summary>
        /// Worksheets the delete.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        public static void WorksheetDelete(this IExcelContext ctx, string name)
        {
            var ctx2 = (ExcelContext)ctx;
            ctx2.WB.Worksheets.Delete(name);
        }

        /// <summary>
        /// Worksheets the get.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        public static void WorksheetGet(this IExcelContext ctx, string name)
        {
            ctx.Flush();
            var ctx2 = (ExcelContext)ctx;
            var worksheets = ctx2.WB.Worksheets;
            ctx2.WS = worksheets[name] ?? worksheets.Add(name);
            ctx.DeltaX = ctx.DeltaY = ctx.XStart = ctx.Y = 1;
        }

        /// <summary>
        /// Worksheets the move.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="name">The name.</param>
        /// <param name="targetName">Name of the target.</param>
        public static void WorksheetMove(this IExcelContext ctx, string name, string targetName)
        {
            ctx.Flush();
            var ctx2 = (ExcelContext)ctx;
            var worksheets = ctx2.WB.Worksheets;
            if (targetName == "<") worksheets.MoveToStart(name);
            else if (targetName == ">") worksheets.MoveToEnd(name);
            else if (targetName.StartsWith("<")) worksheets.MoveBefore(name, targetName.Substring(1));
            else if (targetName.StartsWith(">")) worksheets.MoveAfter(name, targetName.Substring(1));
            else worksheets.MoveAfter(name, targetName);
        }

        // https://www.c-sharpcorner.com/blogs/how-to-adding-pictures-or-images-in-excel-sheet-using-epplus-net-application-c-sharp-part-five
        /// <summary>
        /// Drawings the specified address.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="address">The address.</param>
        /// <param name="name">The name.</param>
        /// <param name="value">The value.</param>
        /// <param name="drawingKind">Kind of the drawing.</param>
        /// <exception cref="ArgumentOutOfRangeException">drawingKind</exception>
        public static void Drawing(this IExcelContext ctx, string address, string name, object value, DrawingKind drawingKind)
        {
            // drawings
            var token = JsonDocument.Parse(value is string @string ? @string : JsonSerializer.Serialize(value)).RootElement;
            var drawings = ((ExcelContext)ctx).WS.Drawings;
            var drawing = ApplyDrawing(name, token, drawings, drawingKind);

            // address
            if (string.IsNullOrEmpty(address))
                return;
            var range = ctx.Get(address);
            if (!token.TryGetProperty("from", out var _))
            {
                drawing.From.Column = range.Start.Column;
                drawing.From.Row = range.Start.Row;
            }
            if (!token.TryGetProperty("to", out var _))
            {
                drawing.To.Column = range.End.Column + 1;
                drawing.To.Row = range.End.Row + 1;
            }
        }

        /// <summary>
        /// Views the action.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="value">The value.</param>
        /// <param name="actionKind">Kind of the action.</param>
        /// <exception cref="ArgumentOutOfRangeException">actionKind</exception>
        public static void ViewAction(this IExcelContext ctx, object value, ViewActionKind actionKind)
        {
            var view = ((ExcelContext)ctx).WS.View;
            switch (actionKind)
            {
                case ViewActionKind.FreezePane: ExcelService.CellToInts((string)value, out var row, out var col); view.FreezePanes(row, col); break;
                case ViewActionKind.SetTabSelected: view.SetTabSelected(); break;
                case ViewActionKind.UnfreezePane: view.UnFreezePanes(); break;
                default: throw new ArgumentOutOfRangeException(nameof(actionKind));
            }
        }

        /// <summary>
        /// Protections the specified value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="value">The value.</param>
        /// <param name="protectionKind">Kind of the protection.</param>
        public static void Protection(this IExcelContext ctx, object value, ProtectionKind protectionKind)
        {
            var protection = ((ExcelContext)ctx).WS.Protection;
            switch (protectionKind)
            {
                case ProtectionKind.AllowFormatRows: protection.AllowFormatRows = value.CastValue<bool>(); break;
                case ProtectionKind.AllowSort: protection.AllowSort = value.CastValue<bool>(); break;
                case ProtectionKind.AllowDeleteRows: protection.AllowDeleteRows = value.CastValue<bool>(); break;
                case ProtectionKind.AllowDeleteColumns: protection.AllowDeleteColumns = value.CastValue<bool>(); break;
                case ProtectionKind.AllowInsertHyperlinks: protection.AllowInsertHyperlinks = value.CastValue<bool>(); break;
                case ProtectionKind.AllowInsertRows: protection.AllowInsertRows = value.CastValue<bool>(); break;
                case ProtectionKind.AllowInsertColumns: protection.AllowInsertColumns = value.CastValue<bool>(); break;
                case ProtectionKind.AllowAutoFilter: protection.AllowAutoFilter = value.CastValue<bool>(); break;
                case ProtectionKind.AllowPivotTables: protection.AllowPivotTables = value.CastValue<bool>(); break;
                case ProtectionKind.AllowFormatCells: protection.AllowFormatCells = value.CastValue<bool>(); break;
                case ProtectionKind.AllowEditScenarios: protection.AllowEditScenarios = value.CastValue<bool>(); break;
                case ProtectionKind.AllowEditObject: protection.AllowEditObject = value.CastValue<bool>(); break;
                case ProtectionKind.AllowSelectUnlockedCells: protection.AllowSelectUnlockedCells = value.CastValue<bool>(); break;
                case ProtectionKind.AllowSelectLockedCells: protection.AllowSelectLockedCells = value.CastValue<bool>(); break;
                case ProtectionKind.IsProtected: protection.IsProtected = value.CastValue<bool>(); break;
                case ProtectionKind.AllowFormatColumns: protection.AllowFormatColumns = value.CastValue<bool>(); break;
                case ProtectionKind.SetPassword: protection.SetPassword(value.CastValue<string>()); break;
                default: throw new ArgumentOutOfRangeException(nameof(protectionKind));
            }
        }

        /// <summary>
        /// Conditionals the formatting.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public static void ConditionalFormatting(this IExcelContext ctx, int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(row, col), value, formattingKind, priority, stopIfTrue);
        /// <summary>
        /// Conditionals the formatting.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public static void ConditionalFormatting(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue);
        /// <summary>
        /// Conditionals the formatting.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public static void ConditionalFormatting(this IExcelContext ctx, Address r, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(r, 0, 0), value, formattingKind, priority, stopIfTrue);
        /// <summary>
        /// Conditionals the formatting.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public static void ConditionalFormatting(this IExcelContext ctx, Address r, int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(r, row, col), value, formattingKind, priority, stopIfTrue);
        /// <summary>
        /// Conditionals the formatting.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        public static void ConditionalFormatting(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue);
        /// <summary>
        /// Conditionals the formatting.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="address">The address.</param>
        /// <param name="value">The value.</param>
        /// <param name="formattingKind">Kind of the formatting.</param>
        /// <param name="priority">The priority.</param>
        /// <param name="stopIfTrue">if set to <c>true</c> [stop if true].</param>
        /// <exception cref="ArgumentNullException">value</exception>
        /// <exception cref="ArgumentOutOfRangeException">formattingKind</exception>
        public static void ConditionalFormatting(this IExcelContext ctx, string address, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue)
        {
            if (value == null)
                throw new ArgumentNullException(nameof(value));
            var token = JsonDocument.Parse(value is string @string ? @string : JsonSerializer.Serialize(value)).RootElement;
            var formatting = ((ExcelContext)ctx).WS.ConditionalFormatting;
            var ruleAddress = new ExcelAddress(ctx.DecodeAddress(address));
            ApplyConditionalFormatting(token, formatting, formattingKind, ruleAddress, priority, stopIfTrue);
        }

        #endregion

        #region Cell

        /// <summary>
        /// Cells the style.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="styles">The styles.</param>
        public static void CellStyle(this IExcelContext ctx, int row, int col, params string[] styles) => CellStyle(ctx, ExcelService.GetAddress(row, col), styles);
        /// <summary>
        /// Cells the style.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="styles">The styles.</param>
        public static void CellStyle(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => CellStyle(ctx, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), styles);
        /// <summary>
        /// Cells the style.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="styles">The styles.</param>
        public static void CellStyle(this IExcelContext ctx, Address r, params string[] styles) => CellStyle(ctx, ExcelService.GetAddress(r, 0, 0), styles);
        /// <summary>
        /// Cells the style.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="styles">The styles.</param>
        public static void CellStyle(this IExcelContext ctx, Address r, int row, int col, params string[] styles) => CellStyle(ctx, ExcelService.GetAddress(r, row, col), styles);
        /// <summary>
        /// Cells the style.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="styles">The styles.</param>
        public static void CellStyle(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => CellStyle(ctx, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), styles);
        /// <summary>
        /// Cells the style.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="cells">The cells.</param>
        /// <param name="styles">The styles.</param>
        public static void CellStyle(this IExcelContext ctx, string cells, string[] styles)
        {
            var range = ctx.Get(cells);
            foreach (var style in styles)
                ApplyStyle(style, range.Style, null);
        }

        /// <summary>
        /// Cells the validation.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="rules">The rules.</param>
        public static void CellValidation(this IExcelContext ctx, CellValidationKind validationKind, int row, int col, params string[] rules) => CellValidation(ctx, validationKind, ExcelService.GetAddress(row, col), rules);
        /// <summary>
        /// Cells the validation.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="rules">The rules.</param>
        public static void CellValidation(this IExcelContext ctx, CellValidationKind validationKind, int fromRow, int fromCol, int toRow, int toCol, params string[] rules) => CellValidation(ctx, validationKind, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), rules);
        /// <summary>
        /// Cells the validation.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="r">The r.</param>
        /// <param name="rules">The rules.</param>
        public static void CellValidation(this IExcelContext ctx, CellValidationKind validationKind, Address r, params string[] rules) => CellValidation(ctx, validationKind, ExcelService.GetAddress(r, 0, 0), rules);
        /// <summary>
        /// Cells the validation.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="rules">The rules.</param>
        public static void CellValidation(this IExcelContext ctx, CellValidationKind validationKind, Address r, int row, int col, params string[] rules) => CellValidation(ctx, validationKind, ExcelService.GetAddress(r, row, col), rules);
        /// <summary>
        /// Cells the validation.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="rules">The rules.</param>
        public static void CellValidation(this IExcelContext ctx, CellValidationKind validationKind, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] rules) => CellValidation(ctx, validationKind, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), rules);
        /// <summary>
        /// Cells the validation.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="validationKind">Kind of the validation.</param>
        /// <param name="cells">The cells.</param>
        /// <param name="rules">The rules.</param>
        /// <exception cref="ArgumentOutOfRangeException">validationKind</exception>
        public static void CellValidation(this IExcelContext ctx, CellValidationKind validationKind, string cells, string[] rules)
        {
            var validations = ((ExcelContext)ctx).WS.DataValidations;
            IExcelDataValidation validation;
            switch (validationKind)
            {
                case CellValidationKind.Find: validation = validations.Find(x => x.Address.Address == cells); break;
                case CellValidationKind.AnyValidation: validation = validations.AddAnyValidation(cells); break;
                case CellValidationKind.CustomValidation: validation = validations.AddCustomValidation(cells); break;
                case CellValidationKind.DateTimeValidation: validation = validations.AddDateTimeValidation(cells); break;
                case CellValidationKind.DecimalValidation: validation = validations.AddDecimalValidation(cells); break;
                case CellValidationKind.IntegerValidation: validation = validations.AddIntegerValidation(cells); break;
                case CellValidationKind.ListValidation: validation = validations.AddListValidation(cells); break;
                case CellValidationKind.TextLengthValidation: validation = validations.AddTextLengthValidation(cells); break;
                case CellValidationKind.TimeValidation: validation = validations.AddTimeValidation(cells); break;
                default: throw new ArgumentOutOfRangeException(nameof(validationKind));
            }
            foreach (var rule in rules)
                ApplyCellValidation(rule, validation);
        }

        /// <summary>
        /// Cells the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public static void CellValue(this IExcelContext ctx, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellValue(ExcelService.GetAddress(row, col), value, valueKind);
        /// <summary>
        /// Cells the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public static void CellValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, valueKind);
        /// <summary>
        /// Cells the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public static void CellValue(this IExcelContext ctx, Address r, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellValue(ExcelService.GetAddress(r, 0, 0), value, valueKind);
        /// <summary>
        /// Cells the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public static void CellValue(this IExcelContext ctx, Address r, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellValue(ExcelService.GetAddress(r, row, col), value, valueKind);
        /// <summary>
        /// Cells the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public static void CellValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, valueKind);
        /// <summary>
        /// Cells the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="cells">The cells.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <exception cref="ArgumentOutOfRangeException">valueKind</exception>
        public static void CellValue(this IExcelContext ctx, string cells, object value, CellValueKind valueKind = CellValueKind.Value)
        {
            var advance = false;
            var range = ctx.Get(cells);
            var values = value == null || !(value is Array array) ? new[] { value } : array;
            foreach (var val in values)
            {
                if (advance) range = ctx.Next(range);
                else advance = true;
                switch (valueKind)
                {
                    case CellValueKind.Value: case CellValueKind.Text: range.Value = val; break;
                    case CellValueKind.AutoFilter: range.AutoFilter = val.CastValue<bool>(); break;
                    case CellValueKind.AutoFitColumns: range.AutoFitColumns(); break;
                    case CellValueKind.Comment: range.Comment.Text = (string)val; break;
                    //case CellValueKind.CommentMore: break;
                    //case CellValueKind.ConditionalFormattingMore: break;
                    case CellValueKind.Copy: var range2 = ((ExcelContext)ctx).WS.Cells[ctx.DecodeAddress((string)val)]; range.Copy(range2); break;
                    case CellValueKind.Formula: range.Formula = (string)val; break;
                    case CellValueKind.FormulaR1C1: range.FormulaR1C1 = (string)val; break;
                    case CellValueKind.Hyperlink: range.Hyperlink = new Uri((string)val); break;
                    case CellValueKind.Merge: range.Merge = val.CastValue<bool>(); break;
                    case CellValueKind.RichText: range.RichText.Add((string)val); break;
                    case CellValueKind.RichTextClear: range.RichText.Clear(); break;
                    case CellValueKind.StyleName: range.StyleName = (string)val; break;
                    default: throw new ArgumentOutOfRangeException(nameof(valueKind));
                }
                if (val is DateTime) range.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
            }
        }

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        public static object GetCellValue(this IExcelContext ctx, int row, int col, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellValue(ExcelService.GetAddress(row, col), valueKind);
        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        public static object GetCellValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), valueKind);
        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        public static object GetCellValue(this IExcelContext ctx, Address r, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellValue(ExcelService.GetAddress(r, 0, 0), valueKind);
        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="row">The row.</param>
        /// <param name="col">The col.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        public static object GetCellValue(this IExcelContext ctx, Address r, int row, int col, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellValue(ExcelService.GetAddress(r, row, col), valueKind);
        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="r">The r.</param>
        /// <param name="fromRow">From row.</param>
        /// <param name="fromCol">From col.</param>
        /// <param name="toRow">To row.</param>
        /// <param name="toCol">To col.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        public static object GetCellValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), valueKind);
        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="cells">The cells.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">valueKind</exception>
        public static object GetCellValue(this IExcelContext ctx, string cells, CellValueKind valueKind = CellValueKind.Value)
        {
            var range = ctx.Get(cells);
            switch (valueKind)
            {
                case CellValueKind.Value: return range.Value;
                case CellValueKind.Text: return range.Text;
                case CellValueKind.AutoFilter: return range.AutoFilter;
                case CellValueKind.Comment: return range.Comment.Text;
                //case CellValueKind.ConditionalFormattingMore: return null;
                case CellValueKind.DataValidation: return range.DataValidation; // get-only
                case CellValueKind.Formula: return range.Formula;
                case CellValueKind.FormulaR1C1: return range.FormulaR1C1;
                case CellValueKind.Hyperlink: return range.Hyperlink;
                case CellValueKind.Merge: return range.Merge;
                case CellValueKind.StyleName: return range.StyleName;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        #endregion

        #region Column

        /// <summary>
        /// Deletes the column.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="column">The column.</param>
        public static void DeleteColumn(this IExcelContext ctx, int column) => ((ExcelContext)ctx).WS.DeleteColumn(column);
        /// <summary>
        /// Deletes the column.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="columnFrom">The column from.</param>
        /// <param name="columns">The columns.</param>
        public static void DeleteColumn(this IExcelContext ctx, int columnFrom, int columns) => ((ExcelContext)ctx).WS.DeleteColumn(columnFrom, columns);

        /// <summary>
        /// Inserts the column.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="columnFrom">The column from.</param>
        /// <param name="columns">The columns.</param>
        public static void InsertColumn(this IExcelContext ctx, int columnFrom, int columns) => ((ExcelContext)ctx).WS.InsertColumn(columnFrom, columns);
        /// <summary>
        /// Inserts the column.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="columnFrom">The column from.</param>
        /// <param name="columns">The columns.</param>
        /// <param name="copyStylesFromColumn">The copy styles from column.</param>
        public static void InsertColumn(this IExcelContext ctx, int columnFrom, int columns, int copyStylesFromColumn) => ((ExcelContext)ctx).WS.InsertColumn(columnFrom, columns, copyStylesFromColumn);

        /// <summary>
        /// Columns the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public static void ColumnValue(this IExcelContext ctx, string col, object value, ColumnValueKind valueKind) => ColumnValue(ctx, ExcelService.ColToInt(col), value, valueKind);
        /// <summary>
        /// Columns the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="col">The col.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <exception cref="ArgumentOutOfRangeException">valueKind</exception>
        public static void ColumnValue(this IExcelContext ctx, int col, object value, ColumnValueKind valueKind)
        {
            var advance = false;
            var column = ((ExcelContext)ctx).WS.Column(col);
            var values = value == null || !(value is Array array) ? new[] { value } : array;
            foreach (var val in values)
            {
                if (advance) column = ctx.Next(column);
                else advance = true;
                switch (valueKind)
                {
                    case ColumnValueKind.AutoFit: column.AutoFit(); break; // set-only
                    case ColumnValueKind.BestFit: column.BestFit = val.CastValue<bool>(); break;
                    case ColumnValueKind.Merged: column.Merged = val.CastValue<bool>(); break;
                    case ColumnValueKind.Width: column.Width = val.CastValue<double>(); break;
                    case ColumnValueKind.TrueWidth: column.SetTrueColumnWidth(val.CastValue<double>()); break; // set-only
                    default: throw new ArgumentOutOfRangeException(nameof(valueKind));
                }
            }
        }

        /// <summary>
        /// Gets the column value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="col">The col.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        public static object GetColumnValue(this IExcelContext ctx, string col, ColumnValueKind valueKind) => GetColumnValue(ctx, ExcelService.ColToInt(col), valueKind);
        /// <summary>
        /// Gets the column value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="col">The col.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">valueKind</exception>
        public static object GetColumnValue(this IExcelContext ctx, int col, ColumnValueKind valueKind)
        {
            var column = ((ExcelContext)ctx).WS.Column(col);
            switch (valueKind)
            {
                case ColumnValueKind.BestFit: return column.BestFit;
                case ColumnValueKind.Merged: return column.Merged;
                case ColumnValueKind.Width: return column.Width;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        #endregion

        #region Row

        /// <summary>
        /// Deletes the row.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        public static void DeleteRow(this IExcelContext ctx, int row) => ((ExcelContext)ctx).WS.DeleteRow(row);
        /// <summary>
        /// Deletes the row.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="rowFrom">The row from.</param>
        /// <param name="rows">The rows.</param>
        public static void DeleteRow(this IExcelContext ctx, int rowFrom, int rows) => ((ExcelContext)ctx).WS.DeleteRow(rowFrom, rows);

        /// <summary>
        /// Inserts the row.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="rowFrom">The row from.</param>
        /// <param name="rows">The rows.</param>
        public static void InsertRow(this IExcelContext ctx, int rowFrom, int rows) => ((ExcelContext)ctx).WS.InsertRow(rowFrom, rows);
        /// <summary>
        /// Inserts the row.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="rowFrom">The row from.</param>
        /// <param name="rows">The rows.</param>
        /// <param name="copyStylesFromRow">The copy styles from row.</param>
        public static void InsertRow(this IExcelContext ctx, int rowFrom, int rows, int copyStylesFromRow) => ((ExcelContext)ctx).WS.InsertRow(rowFrom, rows, copyStylesFromRow);

        /// <summary>
        /// Rows the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        public static void RowValue(this IExcelContext ctx, string row, object value, RowValueKind valueKind) => RowValue(ctx, ExcelService.RowToInt(row), value, valueKind);
        /// <summary>
        /// Rows the value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="value">The value.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <exception cref="ArgumentOutOfRangeException">valueKind</exception>
        public static void RowValue(this IExcelContext ctx, int row, object value, RowValueKind valueKind)
        {
            var advance = false;
            var row_ = ((ExcelContext)ctx).WS.Row(row);
            var values = value == null || !(value is Array array) ? new[] { value } : array;
            foreach (var val in values)
            {
                if (advance) row_ = ctx.Next(row_);
                else advance = true;
                switch (valueKind)
                {
                    case RowValueKind.Collapsed: row_.Collapsed = val.CastValue<bool>(); break;
                    case RowValueKind.CustomHeight: row_.CustomHeight = val.CastValue<bool>(); break;
                    case RowValueKind.Height: row_.Height = val.CastValue<double>(); break;
                    case RowValueKind.Hidden: row_.Hidden = val.CastValue<bool>(); break;
                    case RowValueKind.Merged: row_.Merged = val.CastValue<bool>(); break;
                    case RowValueKind.OutlineLevel: row_.OutlineLevel = val.CastValue<int>(); break;
                    case RowValueKind.PageBreak: row_.PageBreak = val.CastValue<bool>(); break;
                    case RowValueKind.Phonetic: row_.Phonetic = val.CastValue<bool>(); break;
                    case RowValueKind.StyleName: row_.StyleName = val.CastValue<string>(); break;
                    default: throw new ArgumentOutOfRangeException(nameof(valueKind));
                }
            }
        }

        /// <summary>
        /// Gets the row value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        public static object GetRowValue(this IExcelContext ctx, string row, RowValueKind valueKind) => GetRowValue(ctx, ExcelService.RowToInt(row), valueKind);
        /// <summary>
        /// Gets the row value.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="row">The row.</param>
        /// <param name="valueKind">Kind of the value.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentOutOfRangeException">valueKind</exception>
        public static object GetRowValue(this IExcelContext ctx, int row, RowValueKind valueKind)
        {
            var row_ = ((ExcelContext)ctx).WS.Row(row);
            switch (valueKind)
            {
                case RowValueKind.Collapsed: return row_.Collapsed;
                case RowValueKind.CustomHeight: return row_.CustomHeight;
                case RowValueKind.Height: return row_.Height;
                case RowValueKind.Hidden: return row_.Hidden;
                case RowValueKind.Merged: return row_.Merged;
                case RowValueKind.OutlineLevel: return row_.OutlineLevel;
                case RowValueKind.PageBreak: return row_.PageBreak;
                case RowValueKind.Phonetic: return row_.Phonetic;
                case RowValueKind.StyleName: return row_.StyleName;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        #endregion

        #region Parse/Apply

        // SYSTEM

        static T ToStaticEnum<T>(string name, T defaultValue = default) =>
            string.IsNullOrEmpty(name) ? defaultValue :
            (T)typeof(T).GetProperty(name, BindingFlags.Public | BindingFlags.Static)?.GetValue(null);
        static Color ParseColor(string name, Color defaultValue = default) =>
            string.IsNullOrEmpty(name) ? defaultValue :
            name.StartsWith("#") ? ColorTranslator.FromHtml(name) :
            ToStaticEnum<Color>(name);
        //static T ToEnum<T>(string name, T defaultValue = default) => !string.IsNullOrEmpty(name) ? (T)Enum.Parse(typeof(T), name) : defaultValue;
        static T ToEnum<T>(JsonElement name, T defaultValue = default)
            => name.ValueKind == JsonValueKind.String && name.GetString() is string z0 && !string.IsNullOrEmpty(z0) ? (T)Enum.Parse(typeof(T), z0)
            : name.ValueKind == JsonValueKind.Number && name.GetInt32() is int z1 ? (T)(object)z1
            : defaultValue;

        // APPLY

        static void ApplyConditionalFormatting(JsonElement token, ExcelConditionalFormattingCollection formatting, ConditionalFormattingKind formattingKind, ExcelAddress ruleAddress, int? priority, bool stopIfTrue)
        {
            void ApplyColorScale(ExcelConditionalFormattingColorScaleValue val, JsonElement t)
            {
                if (t.TryGetProperty("type", out var z2)) val.Type = ToEnum<eExcelConditionalFormattingValueObjectType>(z2);
                if (t.TryGetProperty("color", out z2)) val.Color = ParseColor(z2.GetString(), Color.White);
                if (t.TryGetProperty("value", out z2)) val.Value = z2.GetDouble();
                if (t.TryGetProperty("formula", out z2)) val.Formula = z2.GetString();
            }
            void ApplyIconDataBar(ExcelConditionalFormattingIconDataBarValue val, JsonElement t)
            {
                if (t.TryGetProperty("type", out var z2)) val.Type = ToEnum<eExcelConditionalFormattingValueObjectType>(z2);
                if (t.TryGetProperty("gte", out z2)) val.GreaterThanOrEqualTo = z2.GetBoolean();
                if (t.TryGetProperty("value", out z2)) val.Value = z2.GetDouble();
                if (t.TryGetProperty("formula", out z2)) val.Formula = z2.GetString();
            }

            IExcelConditionalFormattingWithStdDev stdDev = null;
            IExcelConditionalFormattingWithText text = null;
            IExcelConditionalFormattingWithFormula formula = null;
            IExcelConditionalFormattingWithFormula2 formula2 = null;
            IExcelConditionalFormattingWithRank rank = null;
            IExcelConditionalFormattingRule rule;
            switch (formattingKind)
            {
                case ConditionalFormattingKind.AboveAverage: rule = formatting.AddAboveAverage(ruleAddress); break;
                case ConditionalFormattingKind.AboveOrEqualAverage: rule = formatting.AddAboveOrEqualAverage(ruleAddress); break;
                case ConditionalFormattingKind.AboveStdDev: rule = formatting.AddAboveStdDev(ruleAddress); stdDev = (IExcelConditionalFormattingWithStdDev)rule; break;
                case ConditionalFormattingKind.BeginsWith: rule = formatting.AddBeginsWith(ruleAddress); text = (IExcelConditionalFormattingWithText)rule; break;
                case ConditionalFormattingKind.BelowAverage: rule = formatting.AddBelowAverage(ruleAddress); break;
                case ConditionalFormattingKind.BelowOrEqualAverage: rule = formatting.AddBelowOrEqualAverage(ruleAddress); break;
                case ConditionalFormattingKind.BelowStdDev: rule = formatting.AddBelowStdDev(ruleAddress); stdDev = (IExcelConditionalFormattingWithStdDev)rule; break;
                case ConditionalFormattingKind.Between: rule = formatting.AddBetween(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; formula2 = (IExcelConditionalFormattingWithFormula2)rule; break;
                case ConditionalFormattingKind.Bottom: rule = formatting.AddBottom(ruleAddress); rank = (IExcelConditionalFormattingWithRank)rule; break;
                case ConditionalFormattingKind.BottomPercent: rule = formatting.AddBottomPercent(ruleAddress); rank = (IExcelConditionalFormattingWithRank)rule; break;
                case ConditionalFormattingKind.ContainsBlanks: rule = formatting.AddContainsBlanks(ruleAddress); break;
                case ConditionalFormattingKind.ContainsErrors: rule = formatting.AddContainsErrors(ruleAddress); break;
                case ConditionalFormattingKind.ContainsText: rule = formatting.AddContainsText(ruleAddress); text = (IExcelConditionalFormattingWithText)rule; break;
                case ConditionalFormattingKind.Databar:
                    {
                        var r = formatting.AddDatabar(ruleAddress, token.TryGetProperty("showValue", out var z2) ? ToStaticEnum<Color>(z2.GetString()) : Color.Yellow); rule = r;
                        if (token.TryGetProperty("showValue", out z2)) r.ShowValue = z2.GetBoolean();
                        if (token.TryGetProperty("low", out z2)) ApplyIconDataBar(r.LowValue, z2);
                        if (token.TryGetProperty("high", out z2)) ApplyIconDataBar(r.HighValue, z2);
                    }
                    break;
                case ConditionalFormattingKind.DuplicateValues: rule = formatting.AddDuplicateValues(ruleAddress); break;
                case ConditionalFormattingKind.EndsWith: rule = formatting.AddEndsWith(ruleAddress); text = (IExcelConditionalFormattingWithText)rule; break;
                case ConditionalFormattingKind.Equal: rule = formatting.AddEqual(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.Expression: rule = formatting.AddExpression(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.FiveIconSet:
                    {
                        var r = formatting.AddFiveIconSet(ruleAddress, eExcelconditionalFormatting5IconsSetType.Arrows); rule = r;
                        if (token.TryGetProperty("reverse", out var z2)) r.Reverse = z2.GetBoolean();
                        if (token.TryGetProperty("showValue", out z2)) r.ShowValue = z2.GetBoolean();
                        if (token.TryGetProperty("icon1", out z2)) ApplyIconDataBar(r.Icon1, z2);
                        if (token.TryGetProperty("icon2", out z2)) ApplyIconDataBar(r.Icon2, z2);
                        if (token.TryGetProperty("icon3", out z2)) ApplyIconDataBar(r.Icon3, z2);
                        if (token.TryGetProperty("icon4", out z2)) ApplyIconDataBar(r.Icon4, z2);
                        if (token.TryGetProperty("icon5", out z2)) ApplyIconDataBar(r.Icon5, z2);
                    }
                    break;
                case ConditionalFormattingKind.FourIconSet:
                    {
                        var r = formatting.AddFourIconSet(ruleAddress, eExcelconditionalFormatting4IconsSetType.Arrows); rule = r;
                        if (token.TryGetProperty("reverse", out var z2)) r.Reverse = z2.GetBoolean();
                        if (token.TryGetProperty("showValue", out z2)) r.ShowValue = z2.GetBoolean();
                        if (token.TryGetProperty("icon1", out z2)) ApplyIconDataBar(r.Icon1, z2);
                        if (token.TryGetProperty("icon2", out z2)) ApplyIconDataBar(r.Icon2, z2);
                        if (token.TryGetProperty("icon3", out z2)) ApplyIconDataBar(r.Icon3, z2);
                        if (token.TryGetProperty("icon4", out z2)) ApplyIconDataBar(r.Icon4, z2);
                    }
                    break;
                case ConditionalFormattingKind.GreaterThan: rule = formatting.AddGreaterThan(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.GreaterThanOrEqual: rule = formatting.AddGreaterThanOrEqual(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.Last7Days: rule = formatting.AddLast7Days(ruleAddress); break;
                case ConditionalFormattingKind.LastMonth: rule = formatting.AddLastMonth(ruleAddress); break;
                case ConditionalFormattingKind.LastWeek: rule = formatting.AddLastWeek(ruleAddress); break;
                case ConditionalFormattingKind.LessThan: rule = formatting.AddLessThan(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.LessThanOrEqual: rule = formatting.AddLessThanOrEqual(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.NextMonth: rule = formatting.AddNextMonth(ruleAddress); break;
                case ConditionalFormattingKind.NextWeek: rule = formatting.AddNextWeek(ruleAddress); break;
                case ConditionalFormattingKind.NotBetween: rule = formatting.AddNotBetween(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; formula2 = (IExcelConditionalFormattingWithFormula2)rule; break;
                case ConditionalFormattingKind.NotContainsBlanks: rule = formatting.AddNotContainsBlanks(ruleAddress); break;
                case ConditionalFormattingKind.NotContainsErrors: rule = formatting.AddNotContainsErrors(ruleAddress); break;
                case ConditionalFormattingKind.NotContainsText: rule = formatting.AddNotContainsText(ruleAddress); text = (IExcelConditionalFormattingWithText)rule; break;
                case ConditionalFormattingKind.NotEqual: rule = formatting.AddNotEqual(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.ThisMonth: rule = formatting.AddThisMonth(ruleAddress); break;
                case ConditionalFormattingKind.ThisWeek: rule = formatting.AddThisWeek(ruleAddress); break;
                case ConditionalFormattingKind.ThreeColorScale:
                    {
                        var r = formatting.AddThreeColorScale(ruleAddress); rule = r;
                        if (token.TryGetProperty("low", out var z2)) ApplyColorScale(r.LowValue, z2);
                        if (token.TryGetProperty("high", out z2)) ApplyColorScale(r.HighValue, z2);
                        if (token.TryGetProperty("middle", out z2)) ApplyColorScale(r.MiddleValue, z2);
                    }
                    break;
                case ConditionalFormattingKind.ThreeIconSet:
                    {
                        var r = formatting.AddThreeIconSet(ruleAddress, eExcelconditionalFormatting3IconsSetType.Arrows); rule = r;
                        if (token.TryGetProperty("reverse", out var z2)) r.Reverse = z2.GetBoolean();
                        if (token.TryGetProperty("showValue", out z2)) r.ShowValue = z2.GetBoolean();
                        if (token.TryGetProperty("icon1", out z2)) ApplyIconDataBar(r.Icon1, z2);
                        if (token.TryGetProperty("icon2", out z2)) ApplyIconDataBar(r.Icon2, z2);
                        if (token.TryGetProperty("icon3", out z2)) ApplyIconDataBar(r.Icon3, z2);
                    }
                    break;
                case ConditionalFormattingKind.Today: rule = formatting.AddToday(ruleAddress); break;
                case ConditionalFormattingKind.Tomorrow: rule = formatting.AddTomorrow(ruleAddress); break;
                case ConditionalFormattingKind.Top: rule = formatting.AddTop(ruleAddress); rank = (IExcelConditionalFormattingWithRank)rule; break;
                case ConditionalFormattingKind.TopPercent: rule = formatting.AddTopPercent(ruleAddress); rank = (IExcelConditionalFormattingWithRank)rule; break;
                case ConditionalFormattingKind.TwoColorScale:
                    {
                        var r = formatting.AddTwoColorScale(ruleAddress); rule = r;
                        if (token.TryGetProperty("low", out var z2)) ApplyColorScale(r.LowValue, z2);
                        if (token.TryGetProperty("high", out z2)) ApplyColorScale(r.HighValue, z2);
                    }
                    break;
                case ConditionalFormattingKind.UniqueValues: rule = formatting.AddUniqueValues(ruleAddress); break;
                case ConditionalFormattingKind.Yesterday: rule = formatting.AddYesterday(ruleAddress); break;
                default: throw new ArgumentOutOfRangeException(nameof(formattingKind));
            }
            // CUSTOM
            if (stdDev != null && token.TryGetProperty("stdDev", out var z)) stdDev.StdDev = z.GetUInt16();
            if (text != null && token.TryGetProperty("text", out z)) text.Text = z.GetString();
            if (formula != null && token.TryGetProperty("formula", out z)) formula.Formula = z.GetString();
            if (formula2 != null && token.TryGetProperty("formula2", out z)) formula2.Formula2 = z.GetString();
            if (rank != null && token.TryGetProperty("rank", out z)) rank.Rank = z.GetUInt16();
            // RULE
            if (priority != null) rule.Priority = priority.Value;
            rule.StopIfTrue = stopIfTrue;
            if (token.TryGetProperty("styles", out z))
            {
                var styles =
                    z.ValueKind == JsonValueKind.String ? new[] { z.GetString() } :
                    z.ValueKind == JsonValueKind.Array ? z.EnumerateArray().Select(x => x.GetString()) :
                    throw new ArgumentOutOfRangeException("token.styles", z.ToString());
                foreach (var style in styles)
                    ApplyStyle(style, null, rule.Style);
            }
        }

        static ExcelDrawing ApplyDrawing(string name, JsonElement token, ExcelDrawings drawings, DrawingKind drawingKind)
        {
            // image
            Image ParseImage(JsonElement t) => null;

            // parsing base
            void ApplyDrawingBorder(ExcelDrawingBorder val, JsonElement t)
            {
                if (t.TryGetProperty("fill", out var z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("lineStyle", out z2)) val.LineStyle = ToEnum<eLineStyle>(z2);
                if (t.TryGetProperty("lineCap", out z2)) val.LineCap = ToEnum<eLineCap>(z2);
                if (t.TryGetProperty("width", out z2)) val.Width = z2.GetInt32();
            }
            void ApplyDrawingLineEnd(ExcelDrawingLineEnd val, JsonElement t)
            {
                if (t.TryGetProperty("headEnd", out var z2)) val.HeadEnd = ToEnum<eEndStyle>(z2);
                if (t.TryGetProperty("tailEnd", out z2)) val.TailEnd = ToEnum<eEndStyle>(z2);
                if (t.TryGetProperty("tailEndSizeWidth", out z2)) val.TailEndSizeWidth = ToEnum<eEndSize>(z2);
                if (t.TryGetProperty("tailEndSizeHeight", out z2)) val.TailEndSizeHeight = ToEnum<eEndSize>(z2);
                if (t.TryGetProperty("headEndSizeWidth", out z2)) val.HeadEndSizeWidth = ToEnum<eEndSize>(z2);
                if (t.TryGetProperty("headEndSizeHeight", out z2)) val.HeadEndSizeHeight = ToEnum<eEndSize>(z2);
            }
            void ApplyDrawingFill(ExcelDrawingFill val, JsonElement t)
            {
                if (t.TryGetProperty("orientation", out var z2)) val.Style = ToEnum<eFillStyle>(z2);
                if (t.TryGetProperty("color", out z2)) val.Color = ParseColor(z2.GetString());
                if (t.TryGetProperty("transparancy", out z2)) val.Transparancy = z2.GetInt32();
            }
            void ApplyView3D(ExcelView3D val, JsonElement t)
            {
                if (t.TryGetProperty("perspective", out var z2)) val.Perspective = z2.GetDecimal();
                if (t.TryGetProperty("rotX", out z2)) val.RotX = z2.GetDecimal();
                if (t.TryGetProperty("rotY", out z2)) val.RotY = z2.GetDecimal();
                if (t.TryGetProperty("rightAngleAxes", out z2)) val.RightAngleAxes = z2.GetBoolean();
                if (t.TryGetProperty("depthPercent", out z2)) val.DepthPercent = z2.GetInt32();
                if (t.TryGetProperty("heightPercent", out z2)) val.HeightPercent = z2.GetInt32();
            }
            void ApplyTextFont(ExcelTextFont val, JsonElement t)
            {
                if (t.TryGetProperty("latinFont", out var z2)) val.LatinFont = z2.GetString();
                if (t.TryGetProperty("complexFont", out z2)) val.ComplexFont = z2.GetString();
                if (t.TryGetProperty("bold", out z2)) val.Bold = z2.GetBoolean();
                if (t.TryGetProperty("underLine", out z2)) val.UnderLine = ToEnum<eUnderLineType>(z2);
                if (t.TryGetProperty("underLineColor", out z2)) val.UnderLineColor = ParseColor(z2.GetString());
                if (t.TryGetProperty("italic", out z2)) val.Italic = z2.GetBoolean();
                if (t.TryGetProperty("strike", out z2)) val.Strike = ToEnum<eStrikeType>(z2);
                if (t.TryGetProperty("size", out z2)) val.Size = z2.GetSingle();
                if (t.TryGetProperty("color", out z2)) val.Color = ParseColor(z2.GetString());
            }
            void ApplyParagraphCollection(ExcelParagraphCollection val, JsonElement t)
            {
            }
            ExcelPivotTable ParsePivotTable(JsonElement t)
            {
                return null;
            }

            // parsing drawing
            void ApplyPosition(ExcelDrawing.ExcelPosition val, JsonElement t)
            {
                if (t.TryGetProperty("column", out var z2)) val.Column = z2.GetInt32();
                if (t.TryGetProperty("row", out z2)) val.Row = z2.GetInt32();
                if (t.TryGetProperty("columnOff", out z2)) val.ColumnOff = z2.GetInt32();
                if (t.TryGetProperty("rowOff", out z2)) val.RowOff = z2.GetInt32();
            }
            void ApplyDrawing(ExcelDrawing val, JsonElement t)
            {
                if (t.TryGetProperty("editAs", out var z2)) val.EditAs = ToEnum<eEditAs>(z2);
                if (t.TryGetProperty("name", out z2)) val.Name = z2.GetString();
                if (t.TryGetProperty("from", out z2)) ApplyPosition(val.From, z2);
                if (t.TryGetProperty("to", out z2)) ApplyPosition(val.To, z2);
                if (t.TryGetProperty("print", out z2)) val.Print = z2.GetBoolean();
                if (t.TryGetProperty("locked", out z2)) val.Locked = z2.GetBoolean();
                if (t.TryGetProperty("setPosition", out z2) && z2.GetArrayLength() == 2 && z2.EnumerateArray().Cast<JsonElement>().Select(x => x.GetInt32()) is int[] a0) val.SetPosition(a0[0], a0[1]);
                if (t.TryGetProperty("setPosition", out z2) && z2.GetArrayLength() == 4 && z2.EnumerateArray().Cast<JsonElement>().Select(x => x.GetInt32()) is int[] a1) val.SetPosition(a1[0], a1[1], a1[2], a1[3]);
                if (t.TryGetProperty("setSize", out z2) && z2.GetArrayLength() == 1 && z2.EnumerateArray().Cast<JsonElement>().Select(x => x.GetInt32()) is int[] a2) val.SetSize(a2[0]);
                if (t.TryGetProperty("setSize", out z2) && z2.GetArrayLength() == 2 && z2.EnumerateArray().Cast<JsonElement>().Select(x => x.GetInt32()) is int[] a3) val.SetSize(a3[0], a3[1]);
            }

            // parsing chart
            void ApplyChartTitle(ExcelChartTitle val, JsonElement t)
            {
                if (t.TryGetProperty("text", out var z2)) val.Text = z2.GetString();
                if (t.TryGetProperty("overlay", out z2)) val.Overlay = z2.GetBoolean();
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
                if (t.TryGetProperty("fill", out z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("font", out z2)) ApplyTextFont(val.Font, z2);
                if (t.TryGetProperty("richText", out z2)) ApplyParagraphCollection(val.RichText, z2);
                if (t.TryGetProperty("anchorCtr", out z2)) val.AnchorCtr = z2.GetBoolean();
                if (t.TryGetProperty("anchor", out z2)) val.Anchor = ToEnum<eTextAnchoringType>(z2);
                if (t.TryGetProperty("textVertical", out z2)) val.TextVertical = ToEnum<eTextVerticalType>(z2);
                if (t.TryGetProperty("logRotationBase", out z2)) val.Rotation = z2.GetDouble();
            }
            void ApplyChartAxis(ExcelChartAxis val, JsonElement t)
            {
                if (t.TryGetProperty("font", out var z2)) ApplyTextFont(val.Font, z2);
                if (t.TryGetProperty("orientation", out z2)) val.Orientation = ToEnum<eAxisOrientation>(z2);
                if (t.TryGetProperty("logBase", out z2)) val.LogBase = z2.GetDouble();
                if (t.TryGetProperty("minorTimeUnit", out z2)) val.MinorTimeUnit = ToEnum<eTimeUnit>(z2);
                if (t.TryGetProperty("majorUnit", out z2)) val.MajorUnit = z2.GetDouble();
                if (t.TryGetProperty("maxValue", out z2)) val.MaxValue = z2.GetDouble();
                if (t.TryGetProperty("minValue", out z2)) val.MinValue = z2.GetDouble();
                if (t.TryGetProperty("title", out z2)) ApplyChartTitle(val.Title, z2);
                if (t.TryGetProperty("displayUnit", out z2)) val.DisplayUnit = z2.GetDouble();
                if (t.TryGetProperty("tickLabelPosition", out z2)) val.TickLabelPosition = ToEnum<eTickLabelPosition>(z2);
                if (t.TryGetProperty("deleted", out z2)) val.Deleted = z2.GetBoolean();
                if (t.TryGetProperty("minorGridlines", out z2)) ApplyDrawingBorder(val.MinorGridlines, z2);
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
                if (t.TryGetProperty("fill", out z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("labelPosition", out z2)) val.LabelPosition = ToEnum<eTickLabelPosition>(z2);
                if (t.TryGetProperty("sourceLinked", out z2)) val.SourceLinked = z2.GetBoolean();
                if (t.TryGetProperty("format", out z2)) val.Format = z2.GetString();
                if (t.TryGetProperty("crossesAt", out z2)) val.CrossesAt = z2.GetDouble();
                if (t.TryGetProperty("crossBetween", out z2)) val.CrossBetween = ToEnum<eCrossBetween>(z2);
                if (t.TryGetProperty("crosses", out z2)) val.Crosses = ToEnum<eCrosses>(z2);
                //if (t.TryGetProperty("axisPosition", out z2)) val.AxisPosition = ToEnum<eAxisPosition>(z2);
                if (t.TryGetProperty("minorTickMark", out z2)) val.MinorTickMark = ToEnum<eAxisTickMark>(z2);
                if (t.TryGetProperty("majorTickMark", out z2)) val.MajorTickMark = ToEnum<eAxisTickMark>(z2);
                if (t.TryGetProperty("majorGridlines", out z2)) ApplyDrawingBorder(val.MajorGridlines, z2);
                if (t.TryGetProperty("removeGridlines", out z2)) val.RemoveGridlines();
            }
            void ApplyChartDataTable(ExcelChartDataTable val, JsonElement t)
            {
                if (t.TryGetProperty("showHorizontalBorder", out var z2)) val.ShowHorizontalBorder = z2.GetBoolean();
                if (t.TryGetProperty("showVerticalBorder", out z2)) val.ShowVerticalBorder = z2.GetBoolean();
                if (t.TryGetProperty("showOutline", out z2)) val.ShowOutline = z2.GetBoolean();
                if (t.TryGetProperty("showKeys", out z2)) val.ShowKeys = z2.GetBoolean();
                if (t.TryGetProperty("fill", out z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
                if (t.TryGetProperty("font", out z2)) ApplyTextFont(val.Font, z2);
            }
            void ApplyChartPlotArea(ExcelChartPlotArea val, JsonElement t)
            {
                // if (t.TryGetProperty("chartTypes", out var z2)) ApplyChartTypes(val.ChartTypes, z2);
                if (t.TryGetProperty("dataTable", out var z2)) ApplyChartDataTable(val.DataTable, z2);
                if (t.TryGetProperty("fill", out z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
            }
            void ApplyChartLegend(ExcelChartLegend val, JsonElement t)
            {
                if (t.TryGetProperty("position", out var z2)) val.Position = ToEnum<eLegendPosition>(z2);
                if (t.TryGetProperty("overlay", out z2)) val.Overlay = z2.GetBoolean();
                if (t.TryGetProperty("fill", out z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
                if (t.TryGetProperty("font", out z2)) ApplyTextFont(val.Font, z2);
            }
            void ApplyChartAxis2(ExcelChartAxis[] val, JsonElement t)
            {
            }
            void ApplyChartSeries(ExcelChartSeries val, JsonElement t)
            {
            }
            void ApplyChartXml(XmlDocument val, JsonElement t)
            {
            }
            ExcelDrawing ApplyChart(ExcelChart val, JsonElement t)
            {
                ApplyDrawing(val, t);
                if (t.TryGetProperty("yAxis", out var z2)) ApplyChartAxis(val.YAxis, z2);
                if (t.TryGetProperty("useSecondaryAxis", out z2)) val.UseSecondaryAxis = z2.GetBoolean();
                if (t.TryGetProperty("style", out z2)) val.Style = ToEnum<eChartStyle>(z2);
                if (t.TryGetProperty("roundedCorners", out z2)) val.RoundedCorners = z2.GetBoolean();
                if (t.TryGetProperty("showHiddenData", out z2)) val.ShowHiddenData = z2.GetBoolean();
                if (t.TryGetProperty("displayBlanksAs", out z2)) val.DisplayBlanksAs = ToEnum<eDisplayBlanksAs>(z2);
                if (t.TryGetProperty("plotArea", out z2)) ApplyChartPlotArea(val.PlotArea, z2);
                if (t.TryGetProperty("yAxis", out z2)) ApplyChartAxis(val.XAxis, z2);
                if (t.TryGetProperty("legend", out z2)) ApplyChartLegend(val.Legend, z2);
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
                if (t.TryGetProperty("fill", out z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("view3D", out z2)) ApplyView3D(val.View3D, z2);
                //if (t.TryGetProperty("grouping", out z2)) val.Grouping = ToEnum<eGrouping>(z2.GetString());
                if (t.TryGetProperty("showDataLabelsOverMaximum", out z2)) val.ShowDataLabelsOverMaximum = z2.GetBoolean();
                if (t.TryGetProperty("axis", out z2)) ApplyChartAxis2(val.Axis, z2);
                if (t.TryGetProperty("series", out z2)) ApplyChartSeries(val.Series, z2);
                if (t.TryGetProperty("title", out z2)) ApplyChartTitle(val.Title, z2);
                if (t.TryGetProperty("chartXml", out z2)) ApplyChartXml(val.ChartXml, z2);
                if (t.TryGetProperty("varyColors", out z2)) val.VaryColors = z2.GetBoolean();
                //if (t.TryGetProperty("pivotTableSource", out z2)) ParsePivotTable(val.PivotTableSource, z2);
                return val;
            }

            // parsing picture
            ExcelDrawing ApplyPicture(ExcelPicture val, JsonElement t)
            {
                ApplyDrawing(val, t);
                //if (t.TryGetProperty("image", out var z2)) val.Image = ParseImage(z2.GetString());
                //if (t.TryGetProperty("imageFormat", out z2)) ApplyImageFormat(val.ImageFormat, z2);
                if (t.TryGetProperty("fill", out var z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
                //if (t.TryGetProperty("hyperlink", out z2)) val.Hyperlink = new Uri(z2.GetString());
                return val;
            }

            // parsing shape
            ExcelDrawing ApplyShape(ExcelShape val, JsonElement t)
            {
                ApplyDrawing(val, t);
                //if (t.TryGetProperty("style", out var z2)) val.Style = ToEnum<eShapeStyle>(z2);
                if (t.TryGetProperty("fill", out var z2)) ApplyDrawingFill(val.Fill, z2);
                if (t.TryGetProperty("border", out z2)) ApplyDrawingBorder(val.Border, z2);
                if (t.TryGetProperty("lineEnds", out z2)) ApplyDrawingLineEnd(val.LineEnds, z2);
                if (t.TryGetProperty("font", out z2)) ApplyTextFont(val.Font, z2);
                if (t.TryGetProperty("text", out z2)) val.Text = z2.GetString();
                if (t.TryGetProperty("lockText", out z2)) val.LockText = z2.GetBoolean();
                if (t.TryGetProperty("richText", out z2)) ApplyParagraphCollection(val.RichText, z2);
                if (t.TryGetProperty("textAnchoring", out z2)) val.TextAnchoring = ToEnum<eTextAnchoringType>(z2);
                if (t.TryGetProperty("textAnchoringControl", out z2)) val.TextAnchoringControl = z2.GetBoolean();
                if (t.TryGetProperty("textAlignment", out z2)) val.TextAlignment = ToEnum<eTextAlignment>(z2);
                if (t.TryGetProperty("indent", out z2)) val.Indent = z2.GetInt32();
                if (t.TryGetProperty("textVertical", out z2)) val.TextVertical = ToEnum<eTextVerticalType>(z2);
                return val;
            }

            // drawings
            switch (drawingKind)
            {
                case DrawingKind.AddChart:
                    {
                        var chartType = token.TryGetProperty("type", out var z2) ? ToEnum<eChartType>(z2) : eChartType.Pie;
                        var pivotTableSource = token.TryGetProperty("pivotTableSource", out z2) ? ParsePivotTable(z2) : null;
                        if (pivotTableSource == null) return ApplyChart(drawings.AddChart(name, chartType), token);
                        else return ApplyChart(drawings.AddChart(name, chartType, pivotTableSource), token);
                    }
                case DrawingKind.AddPicture:
                    {
                        var image = token.TryGetProperty("image", out var z2) ? ParseImage(z2) : null;
                        var hyperlink = token.TryGetProperty("hyperlink", out z2) ? new Uri(z2.GetString()) : null;
                        if (hyperlink == null) return ApplyPicture(drawings.AddPicture(name, image), token);
                        else return ApplyPicture(drawings.AddPicture(name, image, hyperlink), token);
                    }
                case DrawingKind.AddShape:
                    {
                        var style = token.TryGetProperty("style", out var z2) ? (eShapeStyle?)ToEnum<eShapeStyle>(z2) : null;
                        if (style == null) return null;
                        return ApplyShape(drawings.AddShape(name, style.Value), token);
                    }
                case DrawingKind.Clear: drawings.Clear(); return null;
                case DrawingKind.Remove: drawings.Remove(name); return null;
                default: throw new ArgumentOutOfRangeException(nameof(drawingKind));
            }
        }

        static void ApplyStyle(string style, ExcelStyle excelStyle, ExcelDxfStyleConditionalFormatting excelDxfStyle)
        {
            string NumberformatPrec(string prec, string defaultPrec) => string.IsNullOrEmpty(prec) ? defaultPrec : $"0.{new string('0', int.Parse(prec))}";

            ExcelVerticalAlignmentFont ParseVerticalAlignmentFont(string value)
            {
                if (char.IsDigit(value[0])) return (ExcelVerticalAlignmentFont)int.Parse(value);
                switch (value.ToLowerInvariant())
                {
                    case "none": return ExcelVerticalAlignmentFont.None;
                    case "baseline": return ExcelVerticalAlignmentFont.Baseline;
                    case "subscript": return ExcelVerticalAlignmentFont.Subscript;
                    case "superscript": return ExcelVerticalAlignmentFont.Superscript;
                    default: throw new ArgumentOutOfRangeException(nameof(value), value);
                }
            }

            ExcelFillStyle ParseFillStyle(string value)
            {
                if (char.IsDigit(value[0])) return (ExcelFillStyle)int.Parse(value);
                switch (value.ToLowerInvariant())
                {
                    default: throw new ArgumentOutOfRangeException(nameof(value), value);
                }
            }

            ExcelBorderStyle ParseBorderStyle(string value)
            {
                if (char.IsDigit(value[0])) return (ExcelBorderStyle)int.Parse(value);
                switch (value.ToLowerInvariant())
                {
                    default: throw new ArgumentOutOfRangeException(nameof(value), value);
                }
            }

            ExcelHorizontalAlignment ParseHorizontalAlignment(string value)
            {
                if (char.IsDigit(value[0])) return (ExcelHorizontalAlignment)int.Parse(value);
                switch (value.ToLowerInvariant())
                {
                    case "general": return ExcelHorizontalAlignment.General;
                    case "left": return ExcelHorizontalAlignment.Left;
                    case "center": return ExcelHorizontalAlignment.Center;
                    case "centercontinuous": return ExcelHorizontalAlignment.CenterContinuous;
                    case "right": return ExcelHorizontalAlignment.Right;
                    case "fill": return ExcelHorizontalAlignment.Fill;
                    case "distributed": return ExcelHorizontalAlignment.Distributed;
                    case "justify": return ExcelHorizontalAlignment.Justify;
                    default: throw new ArgumentOutOfRangeException(nameof(value), value);
                }
            }

            ExcelVerticalAlignment ParseVerticalAlignment(string value)
            {
                if (char.IsDigit(value[0])) return (ExcelVerticalAlignment)int.Parse(value);
                switch (value.ToLowerInvariant())
                {
                    case "top": return ExcelVerticalAlignment.Top;
                    case "center": return ExcelVerticalAlignment.Center;
                    case "bottom": return ExcelVerticalAlignment.Bottom;
                    case "distributed": return ExcelVerticalAlignment.Distributed;
                    case "justify": return ExcelVerticalAlignment.Justify;
                    default: throw new ArgumentOutOfRangeException(nameof(value), value);
                }
            }

            ExcelUnderLineType ParseUnderLineType(string value)
            {
                if (char.IsDigit(value[0])) return (ExcelUnderLineType)int.Parse(value);
                switch (value.ToLowerInvariant())
                {
                    case "none": return ExcelUnderLineType.None;
                    case "single": return ExcelUnderLineType.Single;
                    case "double": return ExcelUnderLineType.Double;
                    case "singleaccounting": return ExcelUnderLineType.SingleAccounting;
                    case "doubleaccounting": return ExcelUnderLineType.DoubleAccounting;
                    default: throw new ArgumentOutOfRangeException(nameof(value), value);
                }
            }

            // https://support.office.com/en-us/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
            // number-format
            if (style.StartsWith("n") && excelStyle != null)
            {
                if (style.StartsWith("n:")) excelStyle.Numberformat.Format = style.Substring(2);
                else if (style.StartsWith("n$")) excelStyle.Numberformat.Format = $"_(\"$\"* #,##{NumberformatPrec(style.Substring(2), "0.00")}_);_(\"$\"* \\(#,##{NumberformatPrec(style.Substring(2), "0.00")}\\);_(\"$\"* \" - \"??_);_(@_)"; // "_-$* #,##{NumberformatPrec(value.Substring(2), "0.00")}_-;-$* #,##{NumberformatPrec(value.Substring(2), "0.00")}_-;_-$* \"-\"??_-;_-@_-";
                else if (style.StartsWith("n%")) excelStyle.Numberformat.Format = $"{NumberformatPrec(style.Substring(2), "0")}%";
                else if (style.StartsWith("n,")) excelStyle.Numberformat.Format = $"_(* #,##{NumberformatPrec(style.Substring(2), "0.00")}_);_(* \\(#,##{NumberformatPrec(style.Substring(2), "0.00")}\\);_(* \"-\"??_);_(@_)";
                else if (style == "nd") excelStyle.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                else throw new InvalidOperationException($"{style} not defined");
            }
            else if (style.StartsWith("n") && excelDxfStyle != null)
            {
                if (style.StartsWith("n:")) excelDxfStyle.NumberFormat.Format = style.Substring(2);
                else if (style.StartsWith("n$")) excelDxfStyle.NumberFormat.Format = $"_(\"$\"* #,##{NumberformatPrec(style.Substring(2), "0.00")}_);_(\"$\"* \\(#,##{NumberformatPrec(style.Substring(2), "0.00")}\\);_(\"$\"* \" - \"??_);_(@_)"; // "_-$* #,##{NumberformatPrec(value.Substring(2), "0.00")}_-;-$* #,##{NumberformatPrec(value.Substring(2), "0.00")}_-;_-$* \"-\"??_-;_-@_-";
                else if (style.StartsWith("n%")) excelDxfStyle.NumberFormat.Format = $"{NumberformatPrec(style.Substring(2), "0")}%";
                else if (style.StartsWith("n,")) excelDxfStyle.NumberFormat.Format = $"_(* #,##{NumberformatPrec(style.Substring(2), "0.00")}_);_(* \\(#,##{NumberformatPrec(style.Substring(2), "0.00")}\\);_(* \"-\"??_);_(@_)";
                else if (style == "nd") excelDxfStyle.NumberFormat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                else throw new InvalidOperationException($"{style} not defined");
            }
            // font
            else if (style.StartsWith("f") && excelStyle != null)
            {
                if (style.StartsWith("f:")) excelStyle.Font.Name = style.Substring(2);
                else if (style.StartsWith("fx")) excelStyle.Font.Size = float.Parse(style.Substring(2));
                else if (style.StartsWith("ff")) excelStyle.Font.Family = int.Parse(style.Substring(2));
                else if (style.StartsWith("fc")) excelStyle.Font.Color.SetColor(ParseColor(style.Substring(2)));
                else if (style.StartsWith("fs")) excelStyle.Font.Scheme = style.Substring(2);
                else if (style == "fB") excelStyle.Font.Bold = true;
                else if (style == "fb") excelStyle.Font.Bold = false;
                else if (style == "fI") excelStyle.Font.Italic = true;
                else if (style == "fi") excelStyle.Font.Italic = false;
                else if (style == "fS") excelStyle.Font.Strike = true;
                else if (style == "fs") excelStyle.Font.Strike = false;
                else if (style == "fU") excelStyle.Font.UnderLine = true;
                else if (style == "fu") excelStyle.Font.UnderLine = false;
                else if (style == "fu:") excelStyle.Font.UnderLineType = ParseUnderLineType(style.Substring(3));
                else if (style.StartsWith("fv")) excelStyle.Font.VerticalAlign = ParseVerticalAlignmentFont(style.Substring(2));
                else throw new InvalidOperationException($"{style} not defined");
            }
            else if (style.StartsWith("f") && excelDxfStyle != null)
            {
                //if (style.StartsWith("f:")) excelDxfStyle.Font.Name = style.Substring(2);
                //else if (style.StartsWith("fx")) excelDxfStyle.Font.Size = float.Parse(style.Substring(2));
                //else if (style.StartsWith("ff")) excelDxfStyle.Font.Family = int.Parse(style.Substring(2));
                //else if (style.StartsWith("fc")) excelDxfStyle.Font.Color = ToDxfColor(style.Substring(2));
                //else if (style.StartsWith("fs")) excelDxfStyle.Font.Scheme = style.Substring(2);
                if (style == "fB") excelDxfStyle.Font.Bold = true;
                else if (style == "fb") excelDxfStyle.Font.Bold = false;
                else if (style == "fI") excelDxfStyle.Font.Italic = true;
                else if (style == "fi") excelDxfStyle.Font.Italic = false;
                else if (style == "fS") excelDxfStyle.Font.Strike = true;
                else if (style == "fs") excelDxfStyle.Font.Strike = false;
                else if (style == "fU") excelDxfStyle.Font.Underline = ExcelUnderLineType.Single;
                else if (style == "fu") excelDxfStyle.Font.Underline = ExcelUnderLineType.None;
                else if (style == "fu:") excelDxfStyle.Font.Underline = ParseUnderLineType(style.Substring(3));
                //else if (style.StartsWith("fv")) excelDxfStyle.Font.VerticalAlign = ParseVerticalAlignmentFont(style.Substring(2));
                else throw new InvalidOperationException($"{style} not defined");
            }
            // fill
            else if (style.StartsWith("l") && excelStyle != null)
            {
                if (style.StartsWith("lc"))
                {
                    if (excelStyle.Fill.PatternType == ExcelFillStyle.None || excelStyle.Fill.PatternType == ExcelFillStyle.Solid) excelStyle.Fill.PatternType = ExcelFillStyle.Solid;
                    excelStyle.Fill.BackgroundColor.SetColor(ParseColor(style.Substring(2)));
                }
                else if (style.StartsWith("lf")) excelStyle.Fill.PatternType = ParseFillStyle(style.Substring(2));
            }
            else if (style.StartsWith("l") && excelDxfStyle != null)
            {
                if (style.StartsWith("lc"))
                {
                    if (excelDxfStyle.Fill.PatternType == ExcelFillStyle.None || excelDxfStyle.Fill.PatternType == ExcelFillStyle.Solid) excelDxfStyle.Fill.PatternType = ExcelFillStyle.Solid;
                    excelDxfStyle.Fill.BackgroundColor.Color = ParseColor(style.Substring(2));
                }
                else if (style.StartsWith("lf")) excelDxfStyle.Fill.PatternType = ParseFillStyle(style.Substring(2));
            }
            // border
            else if (style.StartsWith("b") && excelStyle != null)
            {
                if (style.StartsWith("bl")) excelStyle.Border.Left.Style = ParseBorderStyle(style.Substring(2));
                else if (style.StartsWith("br")) excelStyle.Border.Right.Style = ParseBorderStyle(style.Substring(2));
                else if (style.StartsWith("bt")) excelStyle.Border.Top.Style = ParseBorderStyle(style.Substring(2));
                else if (style.StartsWith("bb")) excelStyle.Border.Bottom.Style = ParseBorderStyle(style.Substring(2));
                else if (style == "bdU") excelStyle.Border.DiagonalUp = true;
                else if (style == "bdu") excelStyle.Border.DiagonalUp = false;
                else if (style == "bdD") excelStyle.Border.DiagonalDown = true;
                else if (style == "bdd") excelStyle.Border.DiagonalDown = false;
                else if (style.StartsWith("bd")) excelStyle.Border.Diagonal.Style = ParseBorderStyle(style.Substring(2));
                else if (style.StartsWith("ba")) excelStyle.Border.BorderAround(ParseBorderStyle(style.Substring(2))); // add color option
                else throw new InvalidOperationException($"{style} not defined");
            }
            else if (style.StartsWith("b") && excelDxfStyle != null)
            {
                if (style.StartsWith("bl")) excelDxfStyle.Border.Left.Style = ParseBorderStyle(style.Substring(2));
                else if (style.StartsWith("br")) excelDxfStyle.Border.Right.Style = ParseBorderStyle(style.Substring(2));
                else if (style.StartsWith("bt")) excelDxfStyle.Border.Top.Style = ParseBorderStyle(style.Substring(2));
                else if (style.StartsWith("bb")) excelDxfStyle.Border.Bottom.Style = ParseBorderStyle(style.Substring(2));
                //else if (style == "bdU") excelDxfStyle.Border.DiagonalUp = true;
                //else if (style == "bdu") excelDxfStyle.Border.DiagonalUp = false;
                //else if (style == "bdD") excelDxfStyle.Border.DiagonalDown = true;
                //else if (style == "bdd") excelDxfStyle.Border.DiagonalDown = false;
                //else if (style.StartsWith("bd")) excelDxfStyle.Border.Diagonal.Style = ParseBorderStyle(style.Substring(2));
                //else if (style.StartsWith("ba")) excelDxfStyle.Border.BorderAround(ParseBorderStyle(style.Substring(2))); // add color option
                else throw new InvalidOperationException($"{style} not defined");
            }
            // horizontal-alignment
            else if (style.StartsWith("ha") && excelStyle != null)
            {
                excelStyle.HorizontalAlignment = ParseHorizontalAlignment(style.Substring(2));
            }
            //else if (style.StartsWith("ha") && excelDxfStyle != null)
            //{
            //    excelDxfStyle.HorizontalAlignment = ParseHorizontalAlignment(style.Substring(2));
            //}
            // vertical-alignment
            else if (style.StartsWith("va") && excelStyle != null)
            {
                excelStyle.VerticalAlignment = ParseVerticalAlignment(style.Substring(2));
            }
            //else if (style.StartsWith("va") && excelDxfStyle != null)
            //{
            //    excelDxfStyle.VerticalAlignment = ParseVerticalAlignment(style.Substring(2));
            //}
            // wrap-text
            else if (style.StartsWith("W") && excelStyle != null) excelStyle.WrapText = true;
            else if (style.StartsWith("w") && excelStyle != null) excelStyle.WrapText = false;
            //else if (style.StartsWith("W") && excelDxfStyle != null) excelDxfStyle.WrapText = true;
            //else if (style.StartsWith("w") && excelDxfStyle != null) excelDxfStyle.WrapText = false;
            else throw new InvalidOperationException($"{style} not defined");
        }

        static void ApplyCellValidation(string rule, IExcelDataValidation validation)
        {
            bool TryParseDataValidationWarningStyle(string value, out ExcelDataValidationWarningStyle style)
            {
                switch (value.ToLowerInvariant())
                {
                    case "undefined": style = ExcelDataValidationWarningStyle.undefined; return true;
                    case "stop": style = ExcelDataValidationWarningStyle.stop; return true;
                    case "warning": style = ExcelDataValidationWarningStyle.warning; return true;
                    case "information": style = ExcelDataValidationWarningStyle.information; return true;
                    default: style = default; return false;
                }
            }

            bool TryParseDataValidationOperator(string value, out ExcelDataValidationOperator op)
            {
                switch (value)
                {
                    case "><": case "..": op = ExcelDataValidationOperator.between; return true;
                    case "<>": case "!.": op = ExcelDataValidationOperator.notBetween; return true;
                    case "=": case "==": op = ExcelDataValidationOperator.equal; return true;
                    case "!=": op = ExcelDataValidationOperator.notEqual; return true;
                    case "<": op = ExcelDataValidationOperator.lessThan; return true;
                    case "<=": op = ExcelDataValidationOperator.lessThanOrEqual; return true;
                    case ">": op = ExcelDataValidationOperator.greaterThan; return true;
                    case ">=": op = ExcelDataValidationOperator.greaterThanOrEqual; return true;
                    default: op = default; return false;
                }
            }
            // base
            if (rule == "_") validation.AllowBlank = true;
            else if (rule == ".") validation.AllowBlank = false;
            else if (rule == "I") validation.ShowInputMessage = true;
            else if (rule == "i") validation.ShowInputMessage = false;
            else if (rule == "E") validation.ShowErrorMessage = true;
            else if (rule == "e") validation.ShowErrorMessage = false;
            else if (rule.StartsWith("et:")) validation.ErrorTitle = rule.Substring(3);
            else if (rule.StartsWith("e:")) validation.Error = rule.Substring(2);
            else if (rule.StartsWith("pt:")) validation.PromptTitle = rule.Substring(3);
            else if (rule.StartsWith("p:")) validation.Prompt = rule.Substring(2);
            // error style
            else if (TryParseDataValidationWarningStyle(rule, out var e)) validation.ErrorStyle = e;
            // operator
            else if (validation is IExcelDataValidationWithOperator o && TryParseDataValidationOperator(rule, out var op)) o.Operator = op;
            // formula
            else if (rule.StartsWith("f:"))
            {
                if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormula> f) f.Formula.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaList> fl) fl.Formula.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaDateTime> fdt) fdt.Formula.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaDecimal> fd) fd.Formula.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaInt> fi) fi.Formula.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaTime> ft) ft.Formula.ExcelFormula = rule.Substring(2);
                else throw new ArgumentOutOfRangeException(nameof(rule), $"{rule} not defined");
            }
            else if (rule.StartsWith("f2:"))
            {
                if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormula> f) f.Formula2.ExcelFormula = rule.Substring(3);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime> fdt) fdt.Formula2.ExcelFormula = rule.Substring(3);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal> fd) fd.Formula2.ExcelFormula = rule.Substring(3);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt> fi) fi.Formula2.ExcelFormula = rule.Substring(3);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime> ft) ft.Formula2.ExcelFormula = rule.Substring(3);
                else throw new ArgumentOutOfRangeException(nameof(rule), $"{rule} not defined");
            }
            // value
            else if (rule.StartsWith("v:"))
            {
                if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaList> fl) { var values = fl.Formula.Values; values.Clear(); foreach (var value in rule.Substring(2).Split('|')) values.Add(value); }
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaDateTime> fdt) fdt.Formula.Value = DateTime.TryParse(rule.Substring(2), out var z) ? (DateTime?)z : null;
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaDecimal> fd) fd.Formula.Value = double.TryParse(rule.Substring(2), out var z) ? (double?)z : null;
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaInt> fi) fi.Formula.Value = int.TryParse(rule.Substring(2), out var z) ? (int?)z : null;
                else if (validation is IExcelDataValidationWithFormula<IExcelDataValidationFormulaTime> ft) ft.Formula.Value = DateTime.TryParse(rule.Substring(2), out var z) ? new ExcelTime { Hour = z.Hour, Minute = z.Minute, Second = z.Second } : null;
                else throw new ArgumentOutOfRangeException(nameof(rule), $"{rule} not defined");
            }
            else if (rule.StartsWith("v2:"))
            {
                if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime> fdt) fdt.Formula2.Value = DateTime.TryParse(rule.Substring(3), out var z) ? (DateTime?)z : null;
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal> fd) fd.Formula2.Value = double.TryParse(rule.Substring(3), out var z) ? (double?)z : null;
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt> fi) fi.Formula2.Value = int.TryParse(rule.Substring(3), out var z) ? (int?)z : null;
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime> ft) ft.Formula2.Value = DateTime.TryParse(rule.Substring(3), out var z) ? new ExcelTime { Hour = z.Hour, Minute = z.Minute, Second = z.Second } : null;
                else throw new ArgumentOutOfRangeException(nameof(rule), $"{rule} not defined");
            }
        }

        #endregion
    }
}
