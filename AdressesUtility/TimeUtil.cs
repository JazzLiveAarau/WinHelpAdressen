using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

namespace AdressesUtility
{

    /// <summary>Time utility functions</summary>
    public static class TimeUtil
    {
        /// <summary>Returns date and time (year_month_day_hour_Minute_Second) as a string</summary>
        public static string YearMonthDayHourMinSec()
        {
            DateTime current_time = DateTime.Now;
            int now_year = current_time.Year;
            int now_month = current_time.Month;
            int now_day = current_time.Day;
            int now_hour = current_time.Hour;
            int now_minute = current_time.Minute;
            int now_second = current_time.Second;

            string time_text = now_year.ToString() + "_" + _IntToString(now_month) + "_" + _IntToString(now_day) + "_" + _IntToString(now_hour) + "_" + _IntToString(now_minute) + "_" + _IntToString(now_second);

            return time_text;
        } // YearMonthDayHourMinSec

        /// <summary>Returns date (year_month_day) as a string</summary>
        public static string YearMonthDay()
        {
            DateTime current_time = DateTime.Now;
            int now_year = current_time.Year;
            int now_month = current_time.Month;
            int now_day = current_time.Day;

            string time_text = "_" + now_year.ToString() + "_" + _IntToString(now_month) + "_" + _IntToString(now_day);

            return time_text;
        } // YearMonthDay

        /// <summary>Returns date and time as a string with a '0' added if input number is less that ten (10)</summary>
        private static string _IntToString(int i_int)
        {
            string time_text = i_int.ToString();

            if (i_int <= 9)
            {
                time_text = "0" + time_text;
            }

            return time_text;
        }
    } // Class TimeUtil
}
