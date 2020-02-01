using System;
using System.Collections.Generic;
using System.Text;

namespace Excelsior
{
    public class Column
    {
        public String Letter { get; set; }
        public override string ToString() => Letter;
        public Dictionary<int, String> Cells { get; set; }

        public String Header
        {
            get
            {
                if (Cells[0] == null)
                    return String.Empty;
                else
                    return Cells[0];
            }
            set
            {
                Cells[0] = value;
            }
        }
        public Column(String letter)
        {
            Letter = letter;
            Cells = new Dictionary<int, String>();
        }

        public List<String> GetValueList()
        {
            // Take each of the cells in the Cells dict and put it into a list, filling in empty values as we go.
            // Dict int indices may not be contiguous, but resultant list indices should be.

            // Counter var to track List index
            int count = 0;

            List<String> Values = new List<String>();

            foreach (KeyValuePair<int, String> kvp in Cells)
            {
                if (kvp.Key > count)
                {
                    // Current key has skipped values since last key (i.e. non-contiguous; i.e. blank cells between this and last one)
                    for (int i = count; i < kvp.Key; i++)
                    {
                        Values.Add(String.Empty);
                    }
                    count = kvp.Key;
                }
                Values.Add(kvp.Value);
                count++;
            }

            return Values;
        }
    }
}
