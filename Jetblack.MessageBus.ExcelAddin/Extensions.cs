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

        public static object[,] ToTable(
            this IDictionary<string, IDictionary<string, object>> dataFrame,
            object[] cols,
            object[] rows,
            bool showColHeaders,
            bool showRowHeaders)
        {
            var filteredRows = cols.Length == 1 && cols[0] is ExcelMissing
                ? dataFrame.Keys.ToArray()
                : cols.Select(x => Optional(x, string.Empty)).ToArray();
            var filteredCols = rows.Length == 1 && rows[0] is ExcelMissing
                ? filteredRows.SelectMany(x => dataFrame[x].Keys).Distinct().ToArray()
                : rows.Select(x => Optional(x, string.Empty)).ToArray();

            var colOffset = showRowHeaders ? 1 : 0;
            var rowOffset = showColHeaders ? 1 : 0;

            var table = new object[filteredRows.Length + rowOffset, filteredCols.Length + colOffset];

            if (showRowHeaders)
            {
                for (int r = 0; r < filteredRows.Length; ++r)
                    table[r + rowOffset, 0] = filteredRows[r];
            }

            if (showColHeaders)
            {
                for (int c = 0; c < filteredCols.Length; ++c)
                    table[0, c + colOffset] = filteredCols[c];
            }

            if (showRowHeaders && showColHeaders)
                table[0, 0] = string.Empty;

            for (int r = 0; r < filteredRows.Length; ++r)
            {
                if (!dataFrame.TryGetValue(filteredRows[r], out var row))
                {
                    for (int c = 0; c < filteredCols.Length; ++c)
                        table[r + rowOffset, c + colOffset] = ExcelMissing.Value;
                }
                else
                {
                    for (int c = 0; c < filteredCols.Length; ++c)
                    {
                        table[r + rowOffset, c + colOffset] = row.TryGetValue(filteredCols[c], out var value)
                            ? value
                            : ExcelMissing.Value;
                    }
                }
            }

            return table;
        }
    }
}
