using System;
using System.IO;

namespace ExcelTrans.Commands
{
    public struct CellStyle : IExcelCommand
    {
        public When When { get; }
        public string Cells { get; private set; }
        public string[] Styles { get; private set; }

        public CellStyle(int row, int col, params string[] styles)
            : this(ExcelService.GetAddress(row, col), styles) { }
        public CellStyle(int fromRow, int fromCol, int toRow, int toCol, params string[] styles)
            : this(ExcelService.GetAddress(fromRow, fromCol, toRow, toCol), styles) { }
        public CellStyle(Address r, params string[] styles)
            : this(ExcelService.GetAddress(r, 0, 0), styles) { }
        public CellStyle(Address r, int row, int col, params string[] styles)
            : this(ExcelService.GetAddress(r, row, col), styles) { }
        public CellStyle(Address r, int fromRow, int fromCol, int toRow, int toCol, params string[] styles)
            : this(ExcelService.GetAddress(r, fromRow, fromCol, toRow, toCol), styles) { }
        public CellStyle(string cells, params string[] styles)
        {
            When = When.Normal;
            Cells = cells ?? throw new ArgumentNullException(nameof(cells));
            Styles = styles ?? throw new ArgumentNullException(nameof(styles));
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            Cells = r.ReadString();
            Styles = new string[r.ReadUInt16()];
            for (var i = 0; i < Styles.Length; i++)
                Styles[i] = r.ReadString();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(Cells);
            w.Write((ushort)Styles.Length);
            for (var i = 0; i < Styles.Length; i++)
                w.Write(Styles[i]);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.CellStyle(Cells, Styles);

        void IExcelCommand.Describe(StringWriter w, int pad) { w.WriteLine($"{new string(' ', pad)}CellStyle[{ExcelService.DescribeAddress(Cells)}]: {string.Join(", ", Styles)}"); }
    }
}