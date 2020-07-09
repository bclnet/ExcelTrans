using System;
using System.IO;
using System.Reflection;

namespace ExcelTrans
{
    public class VbaCode
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Code { get; set; }
        public bool? ReadOnly { get; set; }
        public bool? Private { get; set; }

        public VbaCode() { }
        public VbaCode(Assembly assembly, string codeFile, Func<string, string> codeFunc = null) : this(null, assembly, codeFile, codeFunc) { }
        public VbaCode(string name, Assembly assembly, string codeFile, Func<string, string> codeFunc = null)
        {
            Name = name;
            using (var stream = assembly.GetManifestResourceStream(codeFile))
            using (var reader = new StreamReader(stream))
                Code = reader.ReadToEnd();
            if (codeFunc != null) Code = codeFunc(Code);
        }

        internal VbaCode Read(BinaryReader r)
        {
            Name = r.ReadBoolean() ? r.ReadString() : null;
            Description = r.ReadBoolean() ? r.ReadString() : null;
            Code = r.ReadBoolean() ? r.ReadString() : null;
            ReadOnly = r.ReadBoolean() ? (bool?)r.ReadBoolean() : null;
            Private = r.ReadBoolean() ? (bool?)r.ReadBoolean() : null;
            return this;
        }

        internal VbaCode Write(BinaryWriter w)
        {
            w.Write(Name != null); if (Name != null) w.Write(Name);
            w.Write(Description != null); if (Description != null) w.Write(Description);
            w.Write(Code != null); if (Code != null) w.Write(Code);
            w.Write(ReadOnly != null); if (ReadOnly != null) w.Write(ReadOnly.Value);
            w.Write(Private != null); if (Private != null) w.Write(Private.Value);
            return this;
        }

        public string ProcessCode()
        {
            var w = new StringWriter();
            var r = new StringReader(Code);
            string line;
            while ((line = r.ReadLine()) != null)
                if (line.StartsWith("Option Explicit On", StringComparison.OrdinalIgnoreCase)) { w.Write("Option Explicit"); w.WriteLine(line.Substring(18)); }
                else if (line.StartsWith("#region", StringComparison.OrdinalIgnoreCase)) { w.Write('\''); w.WriteLine(line); }
                else if (line.StartsWith("#end region", StringComparison.OrdinalIgnoreCase)) { w.Write('\''); w.WriteLine(line); }
                else w.WriteLine(line);
            return w.ToString();
        }
    }
}
