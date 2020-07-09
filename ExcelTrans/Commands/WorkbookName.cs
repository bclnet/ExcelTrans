using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct WorkbookName : IExcelCommand
    {
        public When When { get; }
        public string Name { get; private set; }
        public string Cells { get; private set; }
        public WorkbookNameKind NameKind { get; private set; }

        public WorkbookName(string name, int row, int col, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(row, col), nameKind) { }
        public WorkbookName(string name, int fromRow, int fromCol, int toRow, int toCol, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), nameKind) { }
        public WorkbookName(string name, Address r, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(r, 0, 0), nameKind) { }
        public WorkbookName(string name, Address r, int row, int col, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(r, row, col), nameKind) { }
        public WorkbookName(string name, Address r, int fromRow, int fromCol, int toRow, int toCol, WorkbookNameKind nameKind = WorkbookNameKind.Add)
            : this(name, ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), nameKind) { }
        public WorkbookName(string name, string cells, WorkbookNameKind nameKind = WorkbookNameKind.Add)
        {
            When = When.Normal;
            Name = name ?? throw new ArgumentNullException(nameof(cells));
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            NameKind = nameKind;
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Name = r.ReadString();
            Cells = r.ReadString();
            NameKind = (WorkbookNameKind)r.ReadInt32();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Name);
            w.Write(Cells);
            w.Write((int)NameKind);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.WorkbookName(Name, Cells, NameKind);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}WorkbookName[{Name}]: {ExcelService.DescribeAddress(Cells)} - {NameKind}"); }
    }
}