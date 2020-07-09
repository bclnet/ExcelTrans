using System;
using System.Globalization;
using System.IO;

namespace ExcelTrans
{
    public class VbaLibrary
    {
        public string Name { get; set; }
        public LibraryId Libid { get; set; }

        public VbaLibrary() { }
        public VbaLibrary(string name, LibraryId libid)
        {
            Name = name;
            Libid = libid;
        }
        public VbaLibrary(string name,
            Guid guid,
            uint majorVersion,
            uint minorVersion,
            uint type,
            string fullPath,
            string displayName)
        {
            Name = name;
            Libid = new LibraryId(guid, majorVersion, minorVersion, type, fullPath, displayName);
        }

        internal VbaLibrary Read(BinaryReader r)
        {
            Name = r.ReadString();
            Libid = LibraryId.Parse(r.ReadString());
            return this;
        }

        internal VbaLibrary Write(BinaryWriter w)
        {
            w.Write(Name);
            w.Write(Libid.ToString());
            return this;
        }

        // https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/3737ef6e-d819-4186-a5f2-6e258ddf66a5
        public struct LibraryId
        {
            public Guid Guid { get; set; }
            public uint MajorVersion { get; set; }
            public uint MinorVersion { get; set; }
            public uint Type { get; set; } // Lcid?
            public string FullPath { get; set; } // Path;
            public string Description { get; set; } // RegName
            public char ReferenceKind { get; private set; }

            public LibraryId(
                Guid guid,
                uint majorVersion,
                uint minorVersion,
                uint type,
                string fullPath,
                string displayName,
                char referenceKind = '\x0')
            {
                Guid = guid;
                MajorVersion = majorVersion;
                MinorVersion = minorVersion;
                Type = type;
                FullPath = fullPath;
                Description = displayName;
                ReferenceKind = referenceKind != '\x0' ? referenceKind : Environment.OSVersion.Platform == PlatformID.Win32NT ? '\x47' : '\x48';
            }

            public static LibraryId Parse(string libid)
            {
                if (!libid.StartsWith(@"*\"))
                    throw new ArgumentOutOfRangeException(nameof(libid));
                var a = libid.Substring(3).Split('#');
                var b = a[1].Split('.');
                return new LibraryId(
                    referenceKind: libid[2],
                    guid: Guid.Parse(a[0]),
                    majorVersion: uint.Parse(b[0], NumberStyles.HexNumber),
                    minorVersion: uint.Parse(b[1], NumberStyles.HexNumber),
                    type: uint.Parse(a[2], NumberStyles.HexNumber),
                    fullPath: a[3],
                    displayName: a[4]);
            }

            public override string ToString() => $@"*\{ReferenceKind}{{{Guid.ToString().ToUpperInvariant()}}}#{MajorVersion:X}.{MinorVersion:X}#{Type:X}#{FullPath}#{Description}";
        }
    }
}
