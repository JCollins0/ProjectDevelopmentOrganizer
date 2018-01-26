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

        /// <summary>
        /// Debug function to log 
        /// </summary>
        /// <typeparam name="T">type of data</typeparam>
        /// <param name="id">identifier string</param>
        /// <param name="message">message</param>
        public static void Log<T>(string id, T message)
        {
            TextBox log = form_ref.LogBox;
            log.AppendText(id + " ");
            log.AppendText(message.ToString());
            log.AppendText(".\n");
        }

        /// <summary>
        /// makes sure returned value is not null if given data is
        /// </summary>
        /// <param name="dat">data</param>
        /// <returns>'-' if dat is null, dat otherwise</returns>
        public static string AvoidNull(string dat)
        {
            return dat == null ? "-" : dat;
        }

        /// <summary>
        /// Convert String column index to int by using base 26 conversion
        /// </summary>
        /// <param name="column_letter">the column letter</param>
        /// <returns>column number equivalent</returns>
        public static int ConvertColumnLetterToNum(string column_letter)
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

        /// <summary>
        /// Convert column number to String of Letters
        /// </summary>
        /// <param name="column_num">the number of the column</param>
        /// <returns>the Letter equivalent of the column number</returns>
        public static string ConvertNumToColumnLetters(int column_num)
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
            return new string(chars);
        }

        /// <summary>
        /// Formats number with commas
        /// </summary>
        /// <param name="number">the number to format</param>
        /// <returns>formatted number ex) 1000 -> 1,000</returns>
        public static string Format_Int(string number)
        {
            StringBuilder sb_num = new StringBuilder();
            for (int i = 0; i < number.Length; i++)
            {
                if (char.IsDigit(number[i]))
                {
                    sb_num.Append(number[i]);
                }

            }
            string num = sb_num.ToString();
            num = num.Replace(",", "");

            int length = num.Length;
            while (length - 3 > 0)
            {
                num = num.Insert(length - 3, ",");
                length -= 3;
            }
            
            return num;
        }

        /// <summary>
        /// Makes sure a string contains only letters
        /// </summary>
        /// <param name="str_col">the string</param>
        /// <returns>true if str_col contains only letters</returns>
        public static bool IsValidColumnLetter(string str_col)
        {
            foreach(char c in str_col)
            {
                if (!char.IsLetter(c))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Date format MM/dd/yyyy
        /// </summary>
        /// <param name="date">the date to be formatted</param>
        /// <returns>string representation of date</returns>
        public static string ConvertDateToString(DateTime date)
        {
            if (date == DateTime.MinValue)
            {
                return "-";
            }
            return date.ToString("MM/dd/yyyy");
        }

        /// <summary>
        /// Converts double to date
        /// </summary>
        /// <param name="date">the decimal OA representation of a date</param>
        /// <returns>DateTime equivalent to date</returns>
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
        public static string GetColumnRange(int col, long start_row, long end_row)
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
        public static string GetRowRange(long row, int start_col, int end_col)
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
        public static string GetRangeText(int col_start, long row_start, int col_end, long row_end)
        {
            return ConvertNumToColumnLetters(col_start) + row_start + ":" + ConvertNumToColumnLetters(col_end) + row_end;
        }
    }
}
