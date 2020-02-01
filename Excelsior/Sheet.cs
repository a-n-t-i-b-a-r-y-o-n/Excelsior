using System;
using System.Collections.Generic;

namespace Excelsior
{
    public class Sheet
    {
        // User-given name
        public String Name { get; set; }

        // Order within the workbook?
        public int SheetID { get; set; }

        // Relationship ID
        public String rID { get; set; }

        // Internal filename (from .rels file by rID)
        public String Filename { get; set; }

        // Dictonary of columns
        public Dictionary<int, Column> Columns { get; set; }

        public Sheet(String name)
        {
            Name = name;
            Columns = new Dictionary<int, Column>();
        }

        // Return individual cell value or String.Empty
        public String GetCellValue(int col, int row)
        {
            if (Columns[col] != null && Columns[col].Cells[row] != null)
                return Columns[col].Cells[row];
            else
                return String.Empty;
        }
        // Return individual cell value when passed as a letter + number
        // NOTE: Uses 1-based indices like Excel
        public String GetCellValue(String column, int row) => GetCellValue(column.ToColumn(), --row);

        public Column GetColumnByHeader(String Header)
        {
            foreach(KeyValuePair<int, Column> c in Columns)
            {
                if (c.Value.Header.Equals(Header, StringComparison.InvariantCultureIgnoreCase))
                    return c.Value;
            }
            return null;
        }

        public Column GetColumnByHeader(List<String> PossibleHeaders)
        {
            foreach(KeyValuePair<int, Column> c in Columns)
            {
                foreach(String possible in PossibleHeaders)
                {
                    if (c.Value.Header.Equals(possible, StringComparison.InvariantCultureIgnoreCase))
                        return c.Value;
                }
            }
            return null;
        }
    }
}
