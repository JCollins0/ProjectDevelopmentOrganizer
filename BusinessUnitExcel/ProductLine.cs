using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessUnitExcel
{
    class ProductLine
    {
        private const int NUM_COLS_FOR_INDIVIDUAL_PAGES = 1;

        private string product_line_name;
        private Dictionary<string, Project> dictionary_project;

        public ProductLine(string product_line_name)
        {
            this.product_line_name = product_line_name;
            dictionary_project = new Dictionary<string, Project>();
        }

        public void AddProject(string project_name, string project_number)
        {
            Project project = new Project(project_name, project_number);
            if (!dictionary_project.Keys.Contains(project_number))
            {
                dictionary_project.Add(project_number, project);
            }
        }


        public IEnumerable<Project> Projects
        {
            get { return dictionary_project.Values; }
        }

        public Project this[string project_number]
        {
            get
            {
                Project proj;
                dictionary_project.TryGetValue(project_number, out proj);
                return proj;
            }
        }

        internal int CalculateTotal(string design_review_type, string key, string value)
        {
            int sum = 0;
            foreach (Project project in dictionary_project.Values)
            {
                ProjectData dat = project[design_review_type];
                if(dat != null && dat[key] != null && dat[key].ToString() == value)
                {
                    sum++;
                }

            }
            return sum;
        }

        public override int GetHashCode()
        {
            return product_line_name.GetHashCode();
        }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("\n[\n");
            builder.Append(product_line_name);
            builder.Append("\n");
            foreach (Project p in dictionary_project.Values)
            {
                builder.Append(p.ToString());
                builder.Append("\n");
            }
            builder.Append("]\n");
            return builder.ToString();
        }

    }
}
