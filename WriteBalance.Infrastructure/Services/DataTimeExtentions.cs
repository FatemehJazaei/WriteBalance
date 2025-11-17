using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WriteBalance.Infrastructure.Services
{
    public static class DateTimeExtentions
    {

        private static readonly PersianCalendar PersianCalendar = new PersianCalendar();

        public static string ToPersianDate(this DateTime dateTime)
        {

            string month = PersianCalendar.GetMonth(dateTime) > 9 ? PersianCalendar.GetMonth(dateTime).ToString() : $"0{PersianCalendar.GetMonth(dateTime)}";
            string day = PersianCalendar.GetDayOfMonth(dateTime) > 9 ? PersianCalendar.GetDayOfMonth(dateTime).ToString() : $"0{PersianCalendar.GetDayOfMonth(dateTime)}";

            return PersianCalendar.GetYear(dateTime) + month + day;

        }

        public static DateTime ToGeorgianDate(this string persianDate, bool isStart)
        {
            var dateInfo = persianDate.Split('/');
            var year = Convert.ToInt32(dateInfo[0]);
            var month = Convert.ToInt32(dateInfo[1]);
            var day = Convert.ToInt32(dateInfo[2]);

            if (isStart)
                return new DateTime(year, month, day, 0, 0, 0, PersianCalendar);
            else
                return new DateTime(year, month, day, 23, 59, 59, PersianCalendar);

        }


        //تشخیص سال کبیسه
        public static bool IsLeapYear(this DateTime date)
        {
            var persianYear = PersianCalendar.GetYear(date);
            return PersianCalendar.IsLeapYear(persianYear);
        }

        public static bool IsLeapYear(int persianYear)
        {
            return PersianCalendar.IsLeapYear(persianYear);
        }

    }
}
