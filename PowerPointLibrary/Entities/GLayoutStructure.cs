using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLibrary.Entities
{

    /// <summary>
    /// Generated layout structure
    /// 
    /// Actione example : 
    /// <!-- g layout : t 6-3 6-9 -->
    /// </summary>
    public class GLayoutStructure
    {
        public int TitleLines { get; set; }


        public List<Row> Rows { get; set; }


        public class Row
        {
            public List<Bloc> Blocs { get; set; }

            public Row()
            {
                Blocs = new List<Bloc>();
            }

        }

        public class Bloc
        {
            public int Columns { get; set; }
            public int Lines { get; set; }

            public Bloc(int columns, int lines)
            {
                this.Columns = columns;
                this.Lines = lines;
            }

         
        }

    }
}
