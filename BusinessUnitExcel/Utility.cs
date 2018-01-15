using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BusinessUnitExcel
{
    class Utility
    {

        public static BusinessUnitOrganizerForm form_ref;
        // Debug function to log 
        public static void Log<T>(string id, T message)
        {
            TextBox log = form_ref.LogBox;
            log.AppendText(id + " ");
            log.AppendText(message.ToString());
            log.AppendText(".\n");
        }

        public static string AvoidNull(string dat)
        {
            return dat == null ? "-" : dat;
        }

        // Convert String column index to int by using base 26 conversion
        public static int ConvertColumnLetterToNum(String column_letter)
        {
            int column_num = 0;
            char a = 'A';
            for (int i = column_letter.Length - 1; i >= 0; i--)
            {
                int place = column_letter.Length - 1 - i;
                int tspow = (int)Math.Pow(26, place);
                int val = tspow * (column_letter[i] - a + 1);
                column_num += val;
            }


            return column_num;
        }

        // Convert column number to String of Letters
        public static String ConvertNumToColumnLetters(int column_num)
        {
            if (column_num == 0)
            {
                return "A";
            }
            StringBuilder builder = new StringBuilder();

            char a = 'A';
            while (column_num > 0)
            {
                int rem = column_num % 26;

                if (rem == 0)
                {
                    // Add Z
                    builder.Append("Z");
                    column_num = (column_num / 26) - 1;
                }
                else
                {
                    // Add letter
                    builder.Append((char)(a + rem - 1));
                    column_num = (column_num / 26);
                }

            }

            char[] chars = builder.ToString().ToCharArray();
            Array.Reverse(chars);
            return new String(chars);
        }

        public static String Format_Int(String number)
        {
            StringBuilder sb_num = new StringBuilder();
            for (int i = 0; i < number.Length; i++)
            {
                if (char.IsDigit(number[i]))
                {
                    sb_num.Append(number[i]);
                }

            }
            String num = sb_num.ToString();
            num = num.Replace(",", "");

            int length = num.Length;
            while (length - 3 > 0)
            {
                num = num.Insert(length - 3, ",");
                length -= 3;
            }


            return num;
        }

        public static String ConvertDateToString(DateTime date)
        {
            if (date == DateTime.MinValue)
            {
                return "-";
            }
            return date.ToString("MM/dd/yyyy");
        }

        public static DateTime ConvertToDate(double date)
        {
            return DateTime.FromOADate(date);
        }

        /// <summary>
        /// A1:A10
        /// </summary>
        /// <param name="col"></param>
        /// <param name="num_rows"></param>
        /// <returns></returns>
        public static String GetColumnRange(int col, long start_row, long end_row)
        {
            return GetRangeText(col, start_row, col, end_row);
        }

        /// <summary>
        /// A1:Z1
        /// </summary>
        /// <param name="row"></param>
        /// <param name="start_col"></param>
        /// <param name="end_col"></param>
        /// <returns></returns>
        public static String GetRowRange(long row, int start_col, int end_col)
        {
            return GetRangeText(start_col, row, end_col, row);
        }

        /// <summary>
        /// A1:Z10
        /// </summary>
        /// <param name="col_start"></param>
        /// <param name="row_start"></param>
        /// <param name="col_end"></param>
        /// <param name="row_end"></param>
        /// <returns></returns>
        public static String GetRangeText(int col_start, long row_start, int col_end, long row_end)
        {
            return ConvertNumToColumnLetters(col_start) + row_start + ":" + ConvertNumToColumnLetters(col_end) + row_end;
        }
    }
}
