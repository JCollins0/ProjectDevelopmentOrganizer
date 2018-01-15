using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace BusinessUnitExcel
{
    class BusinessSegment
    {

        private string business_segment_name;
        private Dictionary<string, ProductLine> dictionary_product_line;

        public BusinessSegment(string business_segment_name)
        {
            this.business_segment_name = business_segment_name;
            dictionary_product_line = new Dictionary<string,ProductLine>();
        }

        public IEnumerable<ProductLine> ProductLines
        {
            get { return dictionary_product_line.Values; }
        }

        public string BusinessSegmentName 
        {
            get
            {
                return business_segment_name;
            }
            
            set
            {
                this.business_segment_name = value;
            }
        }

        public void AddProductLine(string product_line_name)
        {
            ProductLine product_line = new ProductLine(product_line_name);
            if (!dictionary_product_line.Keys.Contains(product_line_name))
            {
                dictionary_product_line.Add(product_line_name, product_line);
            }
        }

        public ProductLine this[string product_line_name]{
            get
            {
                ProductLine pl;
                dictionary_product_line.TryGetValue(product_line_name, out pl);
                return pl;
            }
        }

        internal int CalculateTotal(string design_review_type, string key, string value)
        {
            int sum = 0;
            foreach (ProductLine product in dictionary_product_line.Values)
            {
                sum += product.CalculateTotal(design_review_type, key, value);
            }
            return sum;
        }

        public override int GetHashCode()
        {
            return business_segment_name.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            return GetHashCode() == obj.GetHashCode();
        }

        public override string ToString()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("\n[\n");
            builder.Append(business_segment_name);
            builder.Append("\n");
            foreach (ProductLine p in dictionary_product_line.Values)
            {
                builder.Append(p.ToString());
                builder.Append("\n");
            }
            builder.Append("]\n");
            return builder.ToString();
        }

        
    }
}
