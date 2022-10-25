using System;
using System.Collections.Generic;
using System.Linq;

using ExcelDna.Integration;

namespace Jetblack.MessageBus.ExcelAddin
{
    internal static class Extensions
    {
        public static IDictionary<string, IDictionary<string, object>> Update(
            this IDictionary<string, IDictionary<string, object>> source,
            IDictionary<string, IDictionary<string, object>> updates)
        {
            foreach (var updatedSeries in updates)
            {
                if (!source.TryGetValue(updatedSeries.Key, out var sourceSeries))
                {
                    sourceSeries = new Dictionary<string, object>();
                    source.Add(updatedSeries.Key, sourceSeries);
                }

                foreach (var item in updatedSeries.Value)
                    sourceSeries[item.Key] = item.Value;
            }

            return source;
        }

        public static string Optional(this object obj, string defaultValue)
        {
            if (obj is string)
                return (string)obj;

            if (obj is ExcelMissing)
                return defaultValue;

            return obj.ToString();
        }

        public static bool Check(this object obj, bool defaultValue)
        {
            if (obj is bool)
                return (bool)obj;

            if (obj is ExcelMissing)
                return defaultValue;

            throw new ArgumentException();
        }

        public static bool IsMissing(this object[] array)
        {
            return array.Length == 1 && array[0] is ExcelMissing;
        }

        public static TValue Get<TKey, TValue>(this IDictionary<TKey, TValue> dict, TKey key, TValue defaultValue)
        {
            return dict.TryGetValue(key, out TValue value) ? value : defaultValue;
        }

        public static object[,] ToTable(
            this IDictionary<string, IDictionary<string, object>> dataFrame,
            object[] colKeys,
            object[] rowKeys,
            bool showColHeaders,
            bool showRowHeaders)
        {
            // Find the selected rows and columns, or all if unspecified.
            var rowHeaders = rowKeys.IsMissing()
                ? dataFrame.Keys.ToArray()
                : rowKeys.Select(x => x.Optional(string.Empty)).ToArray();
            var colHeaders = colKeys.IsMissing()
                ? rowHeaders.SelectMany(x => dataFrame[x].Keys).Distinct().ToArray()
                : colKeys.Select(x => x.Optional(string.Empty)).ToArray();

            // We need an extra row/column if headers are being displayed.
            var colOffset = showRowHeaders ? 1 : 0;
            var rowOffset = showColHeaders ? 1 : 0;

            var table = new object[rowHeaders.Length + rowOffset, colHeaders.Length + colOffset];

            if (showRowHeaders)
            {
                for (int r = 0; r < rowHeaders.Length; ++r)
                    table[r + rowOffset, 0] = rowHeaders[r];
            }

            if (showColHeaders)
            {
                for (int c = 0; c < colHeaders.Length; ++c)
                    table[0, c + colOffset] = colHeaders[c];
            }

            if (showRowHeaders && showColHeaders)
                table[0, 0] = string.Empty; // ExcelMissing.Value displays as 0.

            for (int r = 0; r < rowHeaders.Length; ++r)
            {
                if (!dataFrame.TryGetValue(rowHeaders[r], out var row))
                {
                    for (int c = 0; c < colHeaders.Length; ++c)
                        table[r + rowOffset, c + colOffset] = ExcelMissing.Value;
                }
                else
                {
                    for (int c = 0; c < colHeaders.Length; ++c)
                        table[r + rowOffset, c + colOffset] = row.Get(colHeaders[c], ExcelMissing.Value);
                }
            }

            return table;
        }
    }
}
