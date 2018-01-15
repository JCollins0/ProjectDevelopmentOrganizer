using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessUnitExcel
{
    class ProjectData
    {
        private Dictionary<string, object> data;
        private string decision;
        private DateTime decision_date;
        private DateTime review_date;
        private DateTime gate_2_actual;
        private DateTime gate_3a_actual;
        private DateTime gate_3b_actual;
        private DateTime gate_3c_actual;
        private string gate_comments;
        private string trl;

        private const int NUM_COLS_FOR_INDIVIDUAL_PAGES = 9;

        public ProjectData()
        {
            data = new Dictionary<string, object>();
        }

        public object this[string key]
        {
            get
            {
                object o;
                data.TryGetValue(key, out o);
                if(o == null)
                {
                    return "-";
                }
                return o;
            }

            set
            {
                object o;
                if(!data.TryGetValue(key, out o))
                {
                    data.Add(key, value);
                }
            }
        }


        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("\n\t[\n");
            foreach (string s in data.Keys)
            {
                builder.Append(s + " " + data[s] + "\n");

            }
            builder.Append("\t]\n");
            return builder.ToString();
        }
    }
}
