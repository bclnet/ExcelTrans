using ExcelTrans.Commands;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.DataValidation.Formulas.Contracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace ExcelTrans
{
    public static class ExcelExtensions
    {
        static T ToStaticEnum<T>(string name, T defaultValue = default(T)) =>
            string.IsNullOrEmpty(name) ? defaultValue :
            (T)typeof(T).GetProperty(name, BindingFlags.Public | BindingFlags.Static)?.GetValue(null);
        static Color ToColor(string name, Color defaultValue = default(Color)) =>
            string.IsNullOrEmpty(name) ? defaultValue :
            name.StartsWith("#") ? ColorTranslator.FromHtml(name) :
            ToStaticEnum<Color>(name);
        static T ToEnum<T>(string name, T defaultValue = default) => !string.IsNullOrEmpty(name) ? (T)Enum.Parse(typeof(T), name) : defaultValue;
        static string NumberformatPrec(string prec, string defaultPrec) => string.IsNullOrEmpty(prec) ? defaultPrec : $"0.{new string('0', int.Parse(prec))}";

        #region Execute

        public static object ExecuteCmd(this IExcelContext ctx, IExcelCommand[] cmds, out Action after)
        {
            var frame = ctx.Frame;
            var afterActions = new List<Action>();
            Action action2 = null;
            foreach (var cmd in cmds)
                if (cmd == null) { }
                else if (cmd.When <= When.Before) { cmd.Execute(ctx, ref action2); if (action2 != null) { afterActions.Add(action2); action2 = null; } }
                else afterActions.Add(() => { cmd.Execute(ctx, ref action2); if (action2 != null) { afterActions.Add(action2); action2 = null; } });
            after = afterActions.Count > 0 ? () => { foreach (var action in afterActions) action?.Invoke(); } : (Action)null;
            return frame;
        }

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

        public static void WriteRowFirstSet(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.FirstSet, s, out var after);
        public static void WriteRowFirst(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.First, s, out var after);

        public static void AdvanceRow(this IExcelContext ctx) => ctx.CsvY++;
        public static void WriteRow(this IExcelContext ctx, Collection<string> s)
        {
            var ws = ((ExcelContext)ctx).EnsureWorksheet();
            // execute-row-before
            var cr = ctx.ExecuteRow(When.Before, s, out var after);
            if ((cr & CommandRtn.Continue) == CommandRtn.Continue)
                return;
            //
            ctx.X = ctx.XStart;
            for (var i = 0; i < s.Count; i++)
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
            ctx.ExecuteRow(When.After, s, out var after2);
        }

        public static void WriteRowLast(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.Last, s, out var after);
        public static void WriteRowLastSet(this IExcelContext ctx, Collection<string> s) => ctx.ExecuteRow(When.LastSet, s, out var after);

        #endregion

        #region Worksheet

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

        public static void ConditionalFormatting(this IExcelContext ctx, int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(row, col), value, formattingKind, priority, stopIfTrue);
        public static void ConditionalFormatting(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue);
        public static void ConditionalFormatting(this IExcelContext ctx, Address r, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(r, 0, 0), value, formattingKind, priority, stopIfTrue);
        public static void ConditionalFormatting(this IExcelContext ctx, Address r, int row, int col, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(r, row, col), value, formattingKind, priority, stopIfTrue);
        public static void ConditionalFormatting(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue) => ConditionalFormatting(ctx, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, formattingKind, priority, stopIfTrue);
        public static void ConditionalFormatting(this IExcelContext ctx, string address, object value, ConditionalFormattingKind formattingKind, int? priority, bool stopIfTrue)
        {
            void toColorScale(ExcelConditionalFormattingColorScaleValue val, JToken t)
            {
                if (t == null) return;
                val.Type = ToEnum<eExcelConditionalFormattingValueObjectType>((string)t["type"]);
                val.Color = ToColor((string)t["color"], Color.White);
                val.Value = t["value"].CastValue<double>();
                val.Formula = (string)t["formula"];
            }
            void toIconDataBar(ExcelConditionalFormattingIconDataBarValue val, JToken t)
            {
                if (t == null) return;
                val.Type = ToEnum<eExcelConditionalFormattingValueObjectType>((string)t["type"]);
                val.GreaterThanOrEqualTo = t["gte"].CastValue<bool>();
                val.Value = t["value"].CastValue<double>();
                val.Formula = (string)t["formula"];
            }
            var token = value != null ? JToken.Parse(value is string ? (string)value : JsonConvert.SerializeObject(value)) : null;
            var formatting = ((ExcelContext)ctx).WS.ConditionalFormatting;
            var ruleAddress = new ExcelAddress(ctx.DecodeAddress(address));
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
                        var r = formatting.AddDatabar(ruleAddress, ToStaticEnum<Color>((string)token["color"])); rule = r;
                        r.ShowValue = token["showValue"].CastValue<bool>();
                        toIconDataBar(r.LowValue, token["low"]);
                        toIconDataBar(r.HighValue, token["high"]);
                    }
                    break;
                case ConditionalFormattingKind.DuplicateValues: rule = formatting.AddDuplicateValues(ruleAddress); break;
                case ConditionalFormattingKind.EndsWith: rule = formatting.AddEndsWith(ruleAddress); text = (IExcelConditionalFormattingWithText)rule; break;
                case ConditionalFormattingKind.Equal: rule = formatting.AddEqual(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.Expression: rule = formatting.AddExpression(ruleAddress); formula = (IExcelConditionalFormattingWithFormula)rule; break;
                case ConditionalFormattingKind.FiveIconSet:
                    {
                        var r = formatting.AddFiveIconSet(ruleAddress, eExcelconditionalFormatting5IconsSetType.Arrows); rule = r;
                        r.Reverse = token["reverse"].CastValue<bool>();
                        r.ShowValue = token["showValue"].CastValue<bool>();
                        toIconDataBar(r.Icon1, token["icon1"]);
                        toIconDataBar(r.Icon2, token["icon2"]);
                        toIconDataBar(r.Icon3, token["icon3"]);
                        toIconDataBar(r.Icon4, token["icon4"]);
                        toIconDataBar(r.Icon5, token["icon5"]);
                    }
                    break;
                case ConditionalFormattingKind.FourIconSet:
                    {
                        var r = formatting.AddFourIconSet(ruleAddress, eExcelconditionalFormatting4IconsSetType.Arrows); rule = r;
                        r.Reverse = token["reverse"].CastValue<bool>();
                        r.ShowValue = token["showValue"].CastValue<bool>();
                        toIconDataBar(r.Icon1, token["icon1"]);
                        toIconDataBar(r.Icon2, token["icon2"]);
                        toIconDataBar(r.Icon3, token["icon3"]);
                        toIconDataBar(r.Icon4, token["icon4"]);
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
                        toColorScale(r.LowValue, token["low"]);
                        toColorScale(r.HighValue, token["high"]);
                        toColorScale(r.MiddleValue, token["middle"]);
                    }
                    break;
                case ConditionalFormattingKind.ThreeIconSet:
                    {
                        var r = formatting.AddThreeIconSet(ruleAddress, eExcelconditionalFormatting3IconsSetType.Arrows); rule = r;
                        r.Reverse = token["reverse"].CastValue<bool>();
                        r.ShowValue = token["showValue"].CastValue<bool>();
                        toIconDataBar(r.Icon1, token["icon1"]);
                        toIconDataBar(r.Icon2, token["icon2"]);
                        toIconDataBar(r.Icon3, token["icon3"]);
                    }
                    break;
                case ConditionalFormattingKind.Today: rule = formatting.AddToday(ruleAddress); break;
                case ConditionalFormattingKind.Tomorrow: rule = formatting.AddTomorrow(ruleAddress); break;
                case ConditionalFormattingKind.Top: rule = formatting.AddTop(ruleAddress); rank = (IExcelConditionalFormattingWithRank)rule; break;
                case ConditionalFormattingKind.TopPercent: rule = formatting.AddTopPercent(ruleAddress); rank = (IExcelConditionalFormattingWithRank)rule; break;
                case ConditionalFormattingKind.TwoColorScale:
                    {
                        var r = formatting.AddTwoColorScale(ruleAddress); rule = r;
                        toColorScale(r.LowValue, token["low"]);
                        toColorScale(r.HighValue, token["high"]);
                    }
                    break;
                case ConditionalFormattingKind.UniqueValues: rule = formatting.AddUniqueValues(ruleAddress); break;
                case ConditionalFormattingKind.Yesterday: rule = formatting.AddYesterday(ruleAddress); break;
                default: throw new ArgumentOutOfRangeException(nameof(formattingKind));
            }
            // CUSTOM
            if (stdDev != null) stdDev.StdDev = token["stdDev"].CastValue<ushort>();
            if (text != null) text.Text = (string)token["text"];
            if (formula != null) formula.Formula = (string)token["formula"];
            if (formula2 != null) formula2.Formula2 = (string)token["formula2"];
            if (rank != null) rank.Rank = token["rank"].CastValue<ushort>();
            // RULE
            if (priority != null) rule.Priority = priority.Value;
            rule.StopIfTrue = stopIfTrue;
            var stylesAsToken = token["styles"];
            var styles =
                stylesAsToken == null ? null :
                stylesAsToken.Type == JTokenType.String ? new[] { stylesAsToken.ToObject<string>() } :
                stylesAsToken.Type == JTokenType.Array ? stylesAsToken.ToObject<string[]>() :
                null;
            if (styles != null)
                foreach (var style in styles)
                    ApplyCellStyle(style, null, rule.Style);
        }

        #endregion

        #region Cells

        public static void CellsStyle(this IExcelContext ctx, int row, int col, params string[] styles) => CellsStyle(ctx, ExcelService.GetAddress(row, col), styles);
        public static void CellsStyle(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => CellsStyle(ctx, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, params string[] styles) => CellsStyle(ctx, ExcelService.GetAddress(r, 0, 0), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, int row, int col, params string[] styles) => CellsStyle(ctx, ExcelService.GetAddress(r, row, col), styles);
        public static void CellsStyle(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] styles) => CellsStyle(ctx, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), styles);
        public static void CellsStyle(this IExcelContext ctx, string cells, string[] styles)
        {
            var range = ctx.Get(cells);
            foreach (var style in styles)
                ApplyCellStyle(style, range.Style, null);
        }

        public static void CellsValidation(this IExcelContext ctx, DataValidationKind validationKind, int row, int col, params string[] rules) => CellsValidation(ctx, validationKind, ExcelService.GetAddress(row, col), rules);
        public static void CellsValidation(this IExcelContext ctx, DataValidationKind validationKind, int fromRow, int fromCol, int toRow, int toCol, params string[] rules) => CellsValidation(ctx, validationKind, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), rules);
        public static void CellsValidation(this IExcelContext ctx, DataValidationKind validationKind, Address r, params string[] rules) => CellsValidation(ctx, validationKind, ExcelService.GetAddress(r, 0, 0), rules);
        public static void CellsValidation(this IExcelContext ctx, DataValidationKind validationKind, Address r, int row, int col, params string[] rules) => CellsValidation(ctx, validationKind, ExcelService.GetAddress(r, row, col), rules);
        public static void CellsValidation(this IExcelContext ctx, DataValidationKind validationKind, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] rules) => CellsValidation(ctx, validationKind, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), rules);
        public static void CellsValidation(this IExcelContext ctx, DataValidationKind validationKind, string cells, string[] rules)
        {
            var validations = ((ExcelContext)ctx).WS.DataValidations;
            IExcelDataValidation validation;
            switch (validationKind)
            {
                case DataValidationKind.Find: validation = validations.Find(x => x.Address.Address == cells); break;
                case DataValidationKind.AnyValidation: validation = validations.AddAnyValidation(cells); break;
                case DataValidationKind.CustomValidation: validation = validations.AddCustomValidation(cells); break;
                case DataValidationKind.DateTimeValidation: validation = validations.AddDateTimeValidation(cells); break;
                case DataValidationKind.DecimalValidation: validation = validations.AddDecimalValidation(cells); break;
                case DataValidationKind.IntegerValidation: validation = validations.AddIntegerValidation(cells); break;
                case DataValidationKind.ListValidation: validation = validations.AddListValidation(cells); break;
                case DataValidationKind.TextValidation: validation = validations.AddTextLengthValidation(cells); break;
                case DataValidationKind.TimeValidation: validation = validations.AddTimeValidation(cells); break;
                default: throw new ArgumentOutOfRangeException(nameof(validationKind));
            }
            foreach (var rule in rules)
                ApplyCellValidation(rule, validation);
        }

        public static void CellsValue(this IExcelContext ctx, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(row, col), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, Address r, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(r, 0, 0), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, Address r, int row, int col, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(r, row, col), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, object value, CellValueKind valueKind = CellValueKind.Value) => ctx.CellsValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), value, valueKind);
        public static void CellsValue(this IExcelContext ctx, string cells, object value, CellValueKind valueKind = CellValueKind.Value)
        {
            var range = ctx.Get(cells);
            var values = value == null || !(value is Array array) ? new[] { value } : array;
            foreach (var val in values)
            {
                switch (valueKind)
                {
                    case CellValueKind.Text:
                    case CellValueKind.Value: range.Value = val; break;
                    case CellValueKind.AutoFilter: range.AutoFilter = val.CastValue<bool>(); break;
                    case CellValueKind.AutoFitColumns: range.AutoFitColumns(); break;
                    case CellValueKind.Comment: range.Comment.Text = (string)val; break;
                    case CellValueKind.CommentMore: break;
                    case CellValueKind.ConditionalFormattingMore: break;
                    case CellValueKind.Copy: var range2 = ((ExcelContext)ctx).WS.Cells[ctx.DecodeAddress((string)val)]; range.Copy(range2); break;
                    case CellValueKind.Formula: range.Formula = (string)val; break;
                    case CellValueKind.FormulaR1C1: range.FormulaR1C1 = (string)val; break;
                    case CellValueKind.Hyperlink: range.Hyperlink = new Uri((string)val); break;
                    case CellValueKind.Merge: range.Merge = val.CastValue<bool>(); break;
                    case CellValueKind.RichText: range.RichText.Add((string)val); break;
                    case CellValueKind.RichTextClear: range.RichText.Clear(); break;
                    case CellValueKind.StyleName: range.StyleName = (string)val; break;
                    // validation
                    //case CellValueKind.DataValidation: range.DataValidation = v; break;
                    default: throw new ArgumentOutOfRangeException(nameof(valueKind));
                }
                if (val is DateTime) range.Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                range = ctx.Next(range);
            }
        }

        public static object GetCellsValue(this IExcelContext ctx, int row, int col, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(row, col), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, int fromRow, int fromCol, int toRow, int toCol, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, Address r, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(r, 0, 0), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, Address r, int row, int col, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(r, row, col), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, Address r, int fromRow, int fromCol, int toRow, int toCol, CellValueKind valueKind = CellValueKind.Value) => ctx.GetCellsValue(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), valueKind);
        public static object GetCellsValue(this IExcelContext ctx, string cells, CellValueKind valueKind = CellValueKind.Value)
        {
            var range = ctx.Get(cells);
            switch (valueKind)
            {
                case CellValueKind.Value: return range.Value;
                case CellValueKind.Text: return range.Text;
                case CellValueKind.AutoFilter: return range.AutoFilter;
                case CellValueKind.Comment: return range.Comment.Text;
                case CellValueKind.ConditionalFormattingMore: return null;
                case CellValueKind.Formula: return range.Formula;
                case CellValueKind.FormulaR1C1: return range.FormulaR1C1;
                case CellValueKind.Hyperlink: return range.Hyperlink;
                case CellValueKind.Merge: return range.Merge;
                case CellValueKind.StyleName: return range.StyleName;
                // validation
                case CellValueKind.DataValidation: return range.DataValidation;
                default: throw new ArgumentOutOfRangeException(nameof(valueKind));
            }
        }

        #endregion

        #region Column

        public static void DeleteColumn(this IExcelContext ctx, int column) => ((ExcelContext)ctx).WS.DeleteColumn(column);
        public static void DeleteColumn(this IExcelContext ctx, int columnFrom, int columns) => ((ExcelContext)ctx).WS.DeleteColumn(columnFrom, columns);

        public static void InsertColumn(this IExcelContext ctx, int columnFrom, int columns) => ((ExcelContext)ctx).WS.InsertColumn(columnFrom, columns);
        public static void InsertColumn(this IExcelContext ctx, int columnFrom, int columns, int copyStylesFromColumn) => ((ExcelContext)ctx).WS.InsertColumn(columnFrom, columns, copyStylesFromColumn);

        public static void ColumnValue(this IExcelContext ctx, string col, object value, ColumnValueKind valueKind) => ColumnValue(ctx, ExcelService.ColToInt(col), value, valueKind);
        public static void ColumnValue(this IExcelContext ctx, int col, object value, ColumnValueKind valueKind)
        {
            var column = ((ExcelContext)ctx).WS.Column(col);
            var values = value == null || !(value is Array array) ? new[] { value } : array;
            foreach (var val in values)
            {
                switch (valueKind)
                {
                    case ColumnValueKind.AutoFit: column.AutoFit(); break;
                    case ColumnValueKind.BestFit: column.BestFit = val.CastValue<bool>(); break;
                    case ColumnValueKind.Merged: column.Merged = val.CastValue<bool>(); break;
                    case ColumnValueKind.Width: column.Width = val.CastValue<double>(); break;
                    case ColumnValueKind.TrueWidth: column.SetTrueColumnWidth(val.CastValue<double>()); break;
                    default: throw new ArgumentOutOfRangeException(nameof(valueKind));
                }
                column = ctx.Next(column);
            }
        }

        public static object GetColumnValue(this IExcelContext ctx, string col, ColumnValueKind valueKind) => GetColumnValue(ctx, ExcelService.ColToInt(col), valueKind);
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

        public static void SetTrueColumnWidth(this ExcelColumn column, double width)
        {
            // Deduce what the column width would really get set to.
            var z = width >= (1 + 2 / 3)
                ? Math.Round((Math.Round(7 * (width - 1 / 256), 0) - 5) / 7, 2)
                : Math.Round((Math.Round(12 * (width - 1 / 256), 0) - Math.Round(5 * width, 0)) / 12, 2);

            // How far off? (will be less than 1)
            var errorAmt = width - z;

            // Calculate what amount to tack onto the original amount to result in the closest possible setting.
            var adj = width >= 1 + 2 / 3
                ? Math.Round(7 * errorAmt - 7 / 256, 0) / 7
                : Math.Round(12 * errorAmt - 12 / 256, 0) / 12 + (2 / 12);

            // Set width to a scaled-value that should result in the nearest possible value to the true desired setting.
            if (z > 0)
            {
                column.Width = width + adj;
                return;
            }
            column.Width = 0d;
        }

        #endregion

        #region Row

        public static void DeleteRow(this IExcelContext ctx, int row) => ((ExcelContext)ctx).WS.DeleteRow(row);
        public static void DeleteRow(this IExcelContext ctx, int rowFrom, int rows) => ((ExcelContext)ctx).WS.DeleteRow(rowFrom, rows);

        public static void InsertRow(this IExcelContext ctx, int rowFrom, int rows) => ((ExcelContext)ctx).WS.InsertRow(rowFrom, rows);
        public static void InsertRow(this IExcelContext ctx, int rowFrom, int rows, int copyStylesFromRow) => ((ExcelContext)ctx).WS.InsertRow(rowFrom, rows, copyStylesFromRow);

        public static void RowValue(this IExcelContext ctx, string row, object value, RowValueKind valueKind) => RowValue(ctx, ExcelService.RowToInt(row), value, valueKind);
        public static void RowValue(this IExcelContext ctx, int row, object value, RowValueKind valueKind)
        {
            var row_ = ((ExcelContext)ctx).WS.Row(row);
            var values = value == null || !(value is Array array) ? new[] { value } : array;
            foreach (var val in values)
            {
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
                row_ = ctx.Next(row_);
            }
        }

        public static object GetRowValue(this IExcelContext ctx, string row, RowValueKind valueKind) => GetRowValue(ctx, ExcelService.RowToInt(row), valueKind);
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

        // PARSE

        static ExcelVerticalAlignmentFont ParseVerticalAlignmentFont(string value)
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

        static ExcelFillStyle ParseFillStyle(string value)
        {
            if (char.IsDigit(value[0])) return (ExcelFillStyle)int.Parse(value);
            switch (value.ToLowerInvariant())
            {
                default: throw new ArgumentOutOfRangeException(nameof(value), value);
            }
        }

        static ExcelBorderStyle ParseBorderStyle(string value)
        {
            if (char.IsDigit(value[0])) return (ExcelBorderStyle)int.Parse(value);
            switch (value.ToLowerInvariant())
            {
                default: throw new ArgumentOutOfRangeException(nameof(value), value);
            }
        }

        static ExcelHorizontalAlignment ParseHorizontalAlignment(string value)
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

        static ExcelVerticalAlignment ParseVerticalAlignment(string value)
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

        static bool TryParseDataValidationWarningStyle(string value, out ExcelDataValidationWarningStyle style)
        {
            switch (value.ToLowerInvariant())
            {
                case "undefined": case "null": style = ExcelDataValidationWarningStyle.undefined; return true;
                case "stop": style = ExcelDataValidationWarningStyle.stop; return true;
                case "warning": style = ExcelDataValidationWarningStyle.warning; return true;
                case "information": style = ExcelDataValidationWarningStyle.information; return true;
                default: style = default; return false;
            }
        }

        static bool TryParseDataValidationOperator(string value, out ExcelDataValidationOperator op)
        {
            switch (value)
            {
                case "..": case "><": op = ExcelDataValidationOperator.between; return true;
                case "!.": case "<>": op = ExcelDataValidationOperator.notBetween; return true;
                case "==": op = ExcelDataValidationOperator.equal; return true;
                case "!=": op = ExcelDataValidationOperator.notEqual; return true;
                case "<": op = ExcelDataValidationOperator.lessThan; return true;
                case "<=": op = ExcelDataValidationOperator.lessThanOrEqual; return true;
                case ">": op = ExcelDataValidationOperator.greaterThan; return true;
                case ">=": op = ExcelDataValidationOperator.greaterThanOrEqual; return true;
                default: op = default; return false;
            }
        }

        // APPLY

        public static void ApplyCellStyle(string style, ExcelStyle excelStyle, ExcelDxfStyleConditionalFormatting excelDxfStyle = null)
        {
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
                else if (style.StartsWith("fc:")) excelStyle.Font.Color.SetColor(ToColor(style.Substring(3)));
                else if (style.StartsWith("fs:")) excelStyle.Font.Scheme = style.Substring(2);
                else if (style == "fB") excelStyle.Font.Bold = true;
                else if (style == "fb") excelStyle.Font.Bold = false;
                else if (style == "fI") excelStyle.Font.Italic = true;
                else if (style == "fi") excelStyle.Font.Italic = false;
                else if (style == "fS") excelStyle.Font.Strike = true;
                else if (style == "fs") excelStyle.Font.Strike = false;
                else if (style == "f_") excelStyle.Font.UnderLine = true;
                else if (style == "f!_") excelStyle.Font.UnderLine = false;
                //else if (style == "") excelStyle.Font.UnderLineType = ?;
                else if (style.StartsWith("fv")) excelStyle.Font.VerticalAlign = ParseVerticalAlignmentFont(style.Substring(2));
                else throw new InvalidOperationException($"{style} not defined");
            }
            else if (style.StartsWith("f") && excelDxfStyle != null)
            {
                //if (style.StartsWith("f:")) excelDxfStyle.Font.Name = style.Substring(2);
                //else if (style.StartsWith("fx")) excelDxfStyle.Font.Size = float.Parse(style.Substring(2));
                //else if (style.StartsWith("ff")) excelDxfStyle.Font.Family = int.Parse(style.Substring(2));
                //else if (style.StartsWith("fc:")) excelDxfStyle.Font.Color = ToDxfColor(style.Substring(3));
                //else if (style.StartsWith("fs:")) excelDxfStyle.Font.Scheme = style.Substring(2);
                if (style == "fB") excelDxfStyle.Font.Bold = true;
                else if (style == "fb") excelDxfStyle.Font.Bold = false;
                else if (style == "fI") excelDxfStyle.Font.Italic = true;
                else if (style == "fi") excelDxfStyle.Font.Italic = false;
                else if (style == "fS") excelDxfStyle.Font.Strike = true;
                else if (style == "fs") excelDxfStyle.Font.Strike = false;
                else if (style == "f_") excelDxfStyle.Font.Underline = ExcelUnderLineType.Single;
                else if (style == "f!_") excelDxfStyle.Font.Underline = ExcelUnderLineType.None;
                //else if (style == "") excelDxfStyle.Font.UnderLineType = ?;
                //else if (style.StartsWith("fv")) excelDxfStyle.Font.VerticalAlign = ParseVerticalAlignmentFont(style.Substring(2));
                else throw new InvalidOperationException($"{style} not defined");
            }
            // fill
            else if (style.StartsWith("l") && excelStyle != null)
            {
                if (style.StartsWith("lc:"))
                {
                    if (excelStyle.Fill.PatternType == ExcelFillStyle.None || excelStyle.Fill.PatternType == ExcelFillStyle.Solid) excelStyle.Fill.PatternType = ExcelFillStyle.Solid;
                    excelStyle.Fill.BackgroundColor.SetColor(ToColor(style.Substring(3)));
                }
                else if (style.StartsWith("lf")) excelStyle.Fill.PatternType = ParseFillStyle(style.Substring(2));
            }
            else if (style.StartsWith("l") && excelDxfStyle != null)
            {
                if (style.StartsWith("lc:"))
                {
                    if (excelDxfStyle.Fill.PatternType == ExcelFillStyle.None) excelDxfStyle.Fill.PatternType = ExcelFillStyle.Solid;
                    excelDxfStyle.Fill.BackgroundColor.Color = ToColor(style.Substring(3));
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
                //else if (style.StartsWith("bd")) excelDxfStyle.Border.Diagonal.Style = ParseBorderStyle(style.Substring(2));
                //else if (style == "bdU") excelDxfStyle.Border.DiagonalUp = true;
                //else if (style == "bdu") excelDxfStyle.Border.DiagonalUp = false;
                //else if (style == "bdD") excelDxfStyle.Border.DiagonalDown = true;
                //else if (style == "bdd") excelDxfStyle.Border.DiagonalDown = false;
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

        public static void ApplyCellValidation(string rule, IExcelDataValidation validation)
        {
            // base
            if (TryParseDataValidationWarningStyle(rule, out var style)) validation.ErrorStyle = style;
            else if (rule == "_") validation.AllowBlank = true;
            else if (rule == ".") validation.AllowBlank = false;
            else if (rule == "I") validation.ShowInputMessage = true;
            else if (rule == "i") validation.ShowInputMessage = false;
            else if (rule == "E") validation.ShowErrorMessage = true;
            else if (rule == "e") validation.ShowErrorMessage = false;
            else if (rule.StartsWith("et:")) validation.ErrorTitle = rule.Substring(3);
            else if (rule.StartsWith("e:")) validation.Error = rule.Substring(2);
            else if (rule.StartsWith("pt:")) validation.PromptTitle = rule.Substring(3);
            else if (rule.StartsWith("p:")) validation.Prompt = rule.Substring(2);
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
            // formula2
            else if (rule.StartsWith("f2:"))
            {
                if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormula> f) f.Formula2.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime> fdt) fdt.Formula2.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal> fd) fd.Formula2.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt> fi) fi.Formula2.ExcelFormula = rule.Substring(2);
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime> ft) ft.Formula2.ExcelFormula = rule.Substring(2);
                else throw new ArgumentOutOfRangeException(nameof(rule), $"{rule} not defined");
            }
            // value2
            else if (rule.StartsWith("v2:"))
            {
                if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDateTime> fdt) fdt.Formula2.Value = DateTime.TryParse(rule.Substring(2), out var z) ? (DateTime?)z : null;
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaDecimal> fd) fd.Formula2.Value = double.TryParse(rule.Substring(2), out var z) ? (double?)z : null;
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaInt> fi) fi.Formula2.Value = int.TryParse(rule.Substring(2), out var z) ? (int?)z : null;
                else if (validation is IExcelDataValidationWithFormula2<IExcelDataValidationFormulaTime> ft) ft.Formula2.Value = DateTime.TryParse(rule.Substring(2), out var z) ? new ExcelTime { Hour = z.Hour, Minute = z.Minute, Second = z.Second } : null;
                else throw new ArgumentOutOfRangeException(nameof(rule), $"{rule} not defined");
            }
        }

        #endregion
    }
}
