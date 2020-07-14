using ExcelTrans.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    public class PushSet<T> : IExcelCommand, IExcelSet
    {
        public When When { get; }
        public int TakeY { get; private set; }
        public int SkipX { get; private set; }
        public int SkipY { get; private set; }
        public Func<IExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<T, Collection<string>>>> Group { get; private set; }
        public Func<IExcelContext, object, IExcelCommand[]> Cmds { get; private set; }
        List<Collection<string>> _set;

        public PushSet(Func<IExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<T, Collection<string>>>> group, int takeY = 1, int skipX = 0, int skipY = 0, Func<IExcelContext, IGrouping<T, Collection<string>>, IExcelCommand[]> cmds = null)
        {
            if (cmds == null)
                throw new ArgumentNullException(nameof(cmds));

            When = When.Normal;
            TakeY = takeY;
            SkipX = skipX;
            SkipY = skipY;
            Group = group;
            Cmds = (z, x) => cmds(z, (IGrouping<T, Collection<string>>)x);
        }

        void IExcelCommand.Read(BinaryReader r)
        {
            TakeY = r.ReadInt32();
            SkipX = r.ReadInt32();
            SkipY = r.ReadInt32();
            Group = ExcelSerDes.DecodeFunc<IExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<T, Collection<string>>>>(r);
            Cmds = ExcelSerDes.DecodeFunc<IExcelContext, object, IExcelCommand[]>(r);
            _set = new List<Collection<string>>();
        }

        void IExcelCommand.Write(BinaryWriter w)
        {
            w.Write(TakeY);
            w.Write(SkipX);
            w.Write(SkipY);
            ExcelSerDes.EncodeFunc(w, Group);
            ExcelSerDes.EncodeFunc(w, Cmds);
        }

        void IExcelCommand.Execute(IExcelContext ctx, ref Action after) => ctx.Sets.Push(this);

        void IExcelCommand.Describe(StringWriter w, int pad)
        {
            w.WriteLine($"{new string(' ', pad)}PushSet{(TakeY <= 1 ? null : $"[{TakeY}]")}: {(Group != null ? "[group func]" : null)}");
            if (Group != null)
            {
                var fakeCtx = new ExcelContext();
                var fakeSet = new[] { new Collection<string> { "Fake" } };
                var fakeObj = fakeSet.GroupBy(y => y[0]).FirstOrDefault();
                var cmds = Cmds(fakeCtx, fakeObj);
                ExcelSerDes.DescribeCommands(w, pad, cmds);
            }
        }

        void IExcelSet.Add(Collection<string> s) => _set.Add(s);

        void IExcelSet.Execute(IExcelContext ctx)
        {
            ctx.WriteRowFirstSet(null);
            var takeY = _set.Take(TakeY).ToArray();
            if (Group != null)
                foreach (var g in Group(ctx, _set.Skip(TakeY + SkipY)))
                {
                    ctx.WriteRowFirst(null);
                    var frame = ctx.ExecuteCmd(Cmds(ctx, g), out var action);
                    ctx.CsvY = 0;
                    foreach (var v in takeY)
                    {
                        ctx.CsvY--;
                        ctx.WriteRow(v, SkipX);
                    }
                    ctx.CsvY = 0;
                    foreach (var v in g)
                    {
                        ctx.AdvanceRow();
                        ctx.WriteRow(v, SkipX);
                    }
                    action?.Invoke();
                    ctx.WriteRowLast(null);
                    ctx.Frame = frame;
                }
            ctx.WriteRowLastSet(null);
        }
    }
}