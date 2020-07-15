using System;
using System.IO;
using System.Reflection;

namespace ExcelTrans
{
    /// <summary>
    /// Values for the VbaCodeModule and VbaModule command
    /// </summary>
    public class VbaCode
    {
        /// <summary>
        /// The name of the module
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }
        /// <summary>
        /// A description of the module
        /// </summary>
        /// <value>
        /// The description.
        /// </value>
        public string Description { get; set; }
        /// <summary>
        /// The code without any module level attributes. Can contain function level attributes.
        /// </summary>
        /// <value>
        /// The code.
        /// </value>
        public string Code { get; set; }
        /// <summary>
        /// If the module is readonly
        /// </summary>
        /// <value>
        /// The read only.
        /// </value>
        public bool? ReadOnly { get; set; }
        /// <summary>
        /// If the module is private
        /// </summary>
        /// <value>
        /// The private.
        /// </value>
        public bool? Private { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCode"/> class.
        /// </summary>
        public VbaCode() { }
        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCode"/> class.
        /// </summary>
        /// <param name="assembly">The assembly.</param>
        /// <param name="codeFile">The code file.</param>
        /// <param name="codeFunc">The code function.</param>
        public VbaCode(Assembly assembly, string codeFile, Func<string, string> codeFunc = null) : this(null, assembly, codeFile, codeFunc) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="VbaCode"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="assembly">The assembly.</param>
        /// <param name="codeFile">The code file.</param>
        /// <param name="codeFunc">The code function.</param>
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

        /// <summary>
        /// Processes the code.
        /// </summary>
        /// <returns></returns>
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
