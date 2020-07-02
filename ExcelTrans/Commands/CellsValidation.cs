using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct CellsValidation : IExcelCommand
    {
        public When When { get; }
        public DataValidationKind ValidationKind { get; private set; }
        public string Cells { get; private set; }
        public string[] Rules { get; private set; }

        public CellsValidation(DataValidationKind validationKind, int row, int col, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(row, col), rules) { }
        public CellsValidation(DataValidationKind validationKind, int fromRow, int fromCol, int toRow, int toCol, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), rules) { }
        public CellsValidation(DataValidationKind validationKind, Address r, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(r, 0, 0), rules) { }
        public CellsValidation(DataValidationKind validationKind, Address r, int row, int col, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(r, row, col), rules) { }
        public CellsValidation(DataValidationKind validationKind, Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] rules)
            : this(validationKind, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), rules) { }
        public CellsValidation(DataValidationKind validationKind, string cells, params string[] rules)
        {
            When = When.Normal;
            ValidationKind = validationKind;
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            Rules = rules ?? throw new ArgumentNullException(nameof(rules));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            ValidationKind = (DataValidationKind)r.ReadInt32();
            Rules = new string[r.ReadUInt16()];
            for (var i = 0; i < Rules.Length; i++)
                Rules[i] = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write((int)ValidationKind);
            w.Write((ushort)Rules.Length);
            for (var i = 0; i < Rules.Length; i++)
                w.Write(Rules[i]);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CellsValidation(ValidationKind, Cells, Rules);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CellsValidation.{ValidationKind}[{ExcelService.DescribeAddress(Cells)}]: {string.Join(", ", Rules)}"); }
    }
}