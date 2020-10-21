using ExcelTrans.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace ExcelTrans.Commands
{
    /// <summary>
    /// Pushes a new Set with `group` and `cmds` onto the context stack
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <seealso cref="ExcelTrans.IExcelCommand" />
    /// <seealso cref="ExcelTrans.IExcelSet" />
    public class PushSet<T> : IExcelCommand, IExcelSet
    {
        /// <summary>
        /// Gets the when.
        /// </summary>
        /// <value>
        /// The when.
        /// </value>
        public When When { get; }
        /// <summary>
        /// Gets the take y.
        /// </summary>
        /// <value>
        /// The take y.
        /// </value>
        public int TakeY { get; private set; }
        /// <summary>
        /// Gets the skip x.
        /// </summary>
        /// <value>
        /// The skip x.
        /// </value>
        public int SkipX { get; private set; }
        /// <summary>
        /// Gets the skip y.
        /// </summary>
        /// <value>
        /// The skip y.
        /// </value>
        public int SkipY { get; private set; }
        /// <summary>
        /// Gets the group.
        /// </summary>
        /// <value>
        /// The group.
        /// </value>
        public Func<IExcelContext, IEnumerable<Collection<string>>, IEnumerable<IGrouping<T, Collection<string>>>> Group { get; private set; }
        /// <summary>
        /// Gets the CMDS.
        /// </summary>
        /// <value>
        /// The CMDS.
        /// </value>
        public Func<IExcelContext, object, IExcelCommand[]> Cmds { get; private set; }
        List<Collection<string>> _set;

        /// <summary>
        /// Initializes a new instance of the <see cref="PushSet{T}"/> class.
        /// </summary>
        /// <param name="group">The group.</param>
        /// <param name="takeY">The take y.</param>
        /// <param name="skipX">The skip x.</param>
        /// <param name="skipY">The skip y.</param>
        /// <param name="cmds">The CMDS.</param>
        /// <exception cref="ArgumentNullException">cmds</exception>
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
                var fakeCtx = new ExcelContext(false);
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
                    var frame = ctx.ExecuteCmd(Cmds(ctx, g), out var action);
                    ctx.WriteRowFirst(null);
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