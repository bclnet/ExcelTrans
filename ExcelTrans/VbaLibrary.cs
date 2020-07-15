using System;
using System.Globalization;
using System.IO;

namespace ExcelTrans
{
    /// <summary>
    /// Values for the VbaReference command
    /// </summary>
    public class VbaLibrary
    {
        /// <summary>
        /// The name of the reference
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; set; }
        /// <summary>
        /// LibID For more info check VbaLibrary.LibraryId
        /// </summary>
        /// <value>
        /// The libid.
        /// </value>
        public LibraryId Libid { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="VbaLibrary"/> class.
        /// </summary>
        public VbaLibrary() { }
        /// <summary>
        /// Initializes a new instance of the <see cref="VbaLibrary"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="libid">The libid.</param>
        public VbaLibrary(string name, LibraryId libid)
        {
            Name = name;
            Libid = libid;
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="VbaLibrary"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="guid">The unique identifier.</param>
        /// <param name="majorVersion">The major version.</param>
        /// <param name="minorVersion">The minor version.</param>
        /// <param name="type">The type.</param>
        /// <param name="fullPath">The full path.</param>
        /// <param name="displayName">The display name.</param>
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

        /// <summary>
        /// LibraryId
        /// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/3737ef6e-d819-4186-a5f2-6e258ddf66a5
        /// </summary>
        public struct LibraryId
        {
            /// <summary>
            /// Gets or sets the unique identifier.
            /// </summary>
            /// <value>
            /// The unique identifier.
            /// </value>
            public Guid Guid { get; set; }
            /// <summary>
            /// Gets or sets the major version.
            /// </summary>
            /// <value>
            /// The major version.
            /// </value>
            public uint MajorVersion { get; set; }
            /// <summary>
            /// Gets or sets the minor version.
            /// </summary>
            /// <value>
            /// The minor version.
            /// </value>
            public uint MinorVersion { get; set; }
            /// <summary>
            /// Gets or sets the type.
            /// </summary>
            /// <value>
            /// The type.
            /// </value>
            public uint Type { get; set; } // Lcid?
            /// <summary>
            /// Gets or sets the full path.
            /// </summary>
            /// <value>
            /// The full path.
            /// </value>
            public string FullPath { get; set; } // Path;
            /// <summary>
            /// Gets or sets the description.
            /// </summary>
            /// <value>
            /// The description.
            /// </value>
            public string Description { get; set; } // RegName
            /// <summary>
            /// Gets the kind of the reference.
            /// </summary>
            /// <value>
            /// The kind of the reference.
            /// </value>
            public char ReferenceKind { get; private set; }

            /// <summary>
            /// Initializes a new instance of the <see cref="LibraryId"/> struct.
            /// </summary>
            /// <param name="guid">The unique identifier.</param>
            /// <param name="majorVersion">The major version.</param>
            /// <param name="minorVersion">The minor version.</param>
            /// <param name="type">The type.</param>
            /// <param name="fullPath">The full path.</param>
            /// <param name="displayName">The display name.</param>
            /// <param name="referenceKind">Kind of the reference.</param>
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

            /// <summary>
            /// Parses the specified libid.
            /// </summary>
            /// <param name="libid">The libid.</param>
            /// <returns></returns>
            /// <exception cref="ArgumentOutOfRangeException">libid</exception>
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

            /// <summary>
            /// Converts to string.
            /// </summary>
            /// <returns>
            /// A <see cref="System.String" /> that represents this instance.
            /// </returns>
            public override string ToString() => $@"*\{ReferenceKind}{{{Guid.ToString().ToUpperInvariant()}}}#{MajorVersion:X}.{MinorVersion:X}#{Type:X}#{FullPath}#{Description}";
        }
    }
}
