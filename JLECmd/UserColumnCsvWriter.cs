using System;
using CsvHelper;

namespace JLECmd
{
    /// <summary>
    ///     Writes <typeparamref name="T" /> with the User column first only when --user is used, so the default layout does
    ///     not change. The other columns are the generated ones of <typeparamref name="T" />, whose User member is
    ///     [CsvIgnore].
    /// </summary>
    internal sealed class UserColumnCsvWriter<T> : CsvRecordWriter<T>
    {
        private readonly CsvRecordWriter<T> _columns = CsvRecordWriter.Create<T>();
        private readonly Func<T, string> _user;

        /// <param name="user">Gets the User value of a record, or null for no User column</param>
        public UserColumnCsvWriter(Func<T, string> user)
        {
            _user = user;
        }

        public override int FieldCount => _columns.FieldCount + (_user == null ? 0 : 1);

        public override void WriteHeader(CsvWriter writer)
        {
            if (_user != null)
            {
                WriteColumnName(writer, "User");
            }

            _columns.WriteHeader(writer);
        }

        public override void Write(CsvWriter writer, T record)
        {
            if (_user != null)
            {
                writer.WriteField(_user(record));
            }

            _columns.Write(writer, record);
        }
    }
}
