using System;
using System.Globalization;

namespace ExcelTrans.Services
{
    /// <summary>
    /// Class CsvReaderSettings.
    /// </summary>
    public class CsvReaderSettings
    {
        string _delimiter;

        /// <summary>
        /// Gets or sets the delimiter.
        /// </summary>
        /// <value>The delimiter.</value>
        public string Delimiter
        {
            get
            {
                if (_delimiter == null)
                    try
                    {
                        _delimiter = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
                        if (_delimiter.Length != 1)
                            _delimiter = ",";
                    }
                    catch (Exception) { _delimiter = ","; }
                return _delimiter;
            }
            set => _delimiter = value;
        }
    }
}