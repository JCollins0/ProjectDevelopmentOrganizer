using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessUnitExcel
{
    class Project
    {
        private string project_name;
        private string project_number;
        private Dictionary<string, ProjectData> project_review_types;
        private const int NUM_COLS_FOR_INDIVIDUAL_PAGES = 3;

        public Project(string project_name, string project_number)
        {
            this.project_name = project_name;
            this.project_number = project_number;
            this.project_review_types = new Dictionary<string, ProjectData>();
        }
        
        public void AddProjectData(string design_review_type)
        {
            ProjectData dat;
            if (!project_review_types.TryGetValue(design_review_type, out dat))
            {
                dat = new ProjectData();
                project_review_types.Add(design_review_type, dat);
            }
        }

        internal bool HasDesignReviewType(string design_review_type)
        {
            ProjectData pd;
            if (project_review_types.TryGetValue(design_review_type, out pd))
            {
                pd = null;
                return true;
            }
            return false;
        }
        
        public ProjectData this[string design_review_type]
        {
            get
            {
                ProjectData pd;
                project_review_types.TryGetValue(design_review_type, out pd);
                return pd;
            }
        }

        public override int GetHashCode()
        {
            return project_number.GetHashCode();
        }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("\n[");
            builder.Append(project_number).Append(" ");
            builder.Append(project_name);
            builder.Append("\n");
            foreach (string p in project_review_types.Keys)
            {
                builder.Append(p).Append("={\n");
                builder.Append(project_review_types[p].ToString());
                builder.Append("}\n");
            }
            builder.Append("]\n");
            return builder.ToString();
        }
    }
}
