using System;
using System.Text.RegularExpressions;

namespace Excelsior
{
    // Increment a column (i.e. letter) index
    public static class IndexExtensions
    {
        public static Regex ColumnRegex = new Regex(@"^\D+");
        public static Regex RowRegex = new Regex(@"\d+$");
        public static String IncrementColumn(this String s)
        {
            char c = s[s.Length - 1];
            // Increment the column string for the next iteration
            if (((int)c + 1) % 65 > 25)
            {
                // Recursively iterate, if possible
                if (s.Length - 1 > 0)
                    return IncrementColumn(s.Substring(0, s.Length - 2)) + "A";
                else
                    return "AA";
            }
            else
                return s.Substring(0, s.Length - 1) + (char)((int)c + 1);
        }

        // Translate an Excel-style alpha+numeral (e.g. A1, C4) index to a tuple with Column and Row
        public static Tuple<String, int> ToIndex(this String LetterAndNumber) => new Tuple<string, int>(ColumnRegex.Match(LetterAndNumber).Value, int.Parse(RowRegex.Match(LetterAndNumber).Value));

        public static int ToColumn(this String Letter)
        {
            String alpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            int col = 0;
            for(int i = Letter.Length; i > 0; i--)
                col += alpha.IndexOf(Letter[i - 1]) * ((int)Math.Pow(26, Letter.Length - i));
            return col;
        }
    }
}
