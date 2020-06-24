using System;

namespace ExcelTrans.Services
{
    /// <summary>
    /// CsvWriterColumn
    /// </summary>
    public class CsvWriterColumn
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CsvWriterColumn"/> class.
        /// </summary>
        protected CsvWriterColumn() { }
        /// <summary>
        /// Initializes a new instance of the <see cref="CsvWriterColumn"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="getValue">The get value.</param>
        /// <param name="displayName">The display name.</param>
        public CsvWriterColumn(string name, Func<object, object> getValue, string displayName = null)
        {
            Name = name;
            GetValue = getValue;
            DisplayName = displayName;
        }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        public string Name { get; protected set; }

        /// <summary>
        /// Gets or sets the get value.
        /// </summary>
        /// <value>
        /// The get value.
        /// </value>
        public Func<object, object> GetValue { get; protected set; }

        /// <summary>
        /// Gets or sets the display name.
        /// </summary>
        /// <value>
        /// The display name.
        /// </value>
        public string DisplayName { get; protected set; }
    }

    /// <summary>
    /// CsvWriterColumn
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class CsvWriterColumn<T> : CsvWriterColumn
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CsvWriterColumn{T}"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <param name="getValue">The get value.</param>
        /// <param name="displayName">The display name.</param>
        public CsvWriterColumn(string name, Func<T, object> getValue, string displayName = null)
        {
            Name = name;
            GetValue = x => getValue((T)x);
            DisplayName = displayName;
        }
    }
}