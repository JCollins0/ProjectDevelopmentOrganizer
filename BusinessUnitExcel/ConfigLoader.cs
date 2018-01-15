using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace BusinessUnitExcel
{
    class ConfigLoader
    {
        public static Dictionary<string, List<string>> drtheaders = new Dictionary<string, List<string>>();
        public static List<string> list_drt = new List<string>();
        public static Dictionary<string, int> headerinfo = new Dictionary<string, int>();
        public static List<string> black_listed_columns = new List<string>();
        public static List<string> summary_headers = new List<string>();
        public static List<string> summary_foreach_headers = new List<string>();

        public static List<string> totals_headers = new List<string>();
        public static List<string> totals_foreach_headers = new List<string>();

        private ConfigLoader() { }

        public static bool LoadConfig()
        {
            try
            {
                XElement document = XElement.Load("Config.xml");
                ParseMainFields(document);
                Utility.Log("Info:","Config file found");
                return true;
            }
            catch (Exception)
            {
                Utility.Log("Error:", "Config.xml not found, restart program with Config file in same directory");
                return false;
            }
        }

        private static void ParseMainFields(XElement document)
        {
            foreach (XElement element in document.Elements("DesignReviewType"))
            {
                ParseDesignReviewType(element);
            }

            foreach (XElement element in document.Elements("HeaderInfo"))
            {
                ParseHeaderInfo(element);
            }

            ParseSummary(document.Element("Summary"));
            ParseTotals(document.Element("Totals"));
        }

        private static void ParseTotals(XElement totals)
        {
            foreach (XElement element in totals.Elements("Header"))
            {
                totals_headers.Add(element.Value);
            }

            foreach (XElement element in totals.Element("ForEach").Elements("Header"))
            {
                totals_foreach_headers.Add(element.Value);
            }
        }

        private static void ParseSummary(XElement summary)
        {
            foreach (XElement element in summary.Elements("Header"))
            {
                summary_headers.Add(element.Value);
            }

            foreach (XElement element in summary.Element("ForEach").Elements("Header"))
            {
                summary_foreach_headers.Add(element.Value);
            }

        }

        private static void ParseHeaderInfo(XElement element)
        {
            
            string key = element.Element("Header").Value;
            string value = element.Element("Column").Value;

            headerinfo.Add(key, Utility.ConvertColumnLetterToNum(value));
            black_listed_columns.Add(value);
        }

        private static void ParseDesignReviewType(XElement element)
        {
            string name = element.Element("Name").Value;

            list_drt.Add(name);

            List<string> headers = new List<string>();

            foreach (XElement header in element.Elements("Header"))
            {
                headers.Add(header.Value);
            }

            drtheaders.Add(name, headers);
        }

    }
}
