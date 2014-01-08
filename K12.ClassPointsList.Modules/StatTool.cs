using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;

namespace K12ClassPointsList.Modules
{
    static class StatTool
    {
        private static QueryHelper _q;
        /// <summary>
        /// QueryHelper
        /// </summary>
        public static QueryHelper Q
        {
            get
            {
                if (_q == null)
                    _q = new QueryHelper();
                return _q;
            }
        }

        public static string CheckWeek(string x)
        {
            #region 依編號取代為星期
            if (x == "Monday")
            {
                return "一";
            }
            else if (x == "Tuesday")
            {
                return "二";
            }
            else if (x == "Wednesday")
            {
                return "三";
            }
            else if (x == "Thursday")
            {
                return "四";
            }
            else if (x == "Friday")
            {
                return "五";
            }
            else if (x == "Saturday")
            {
                return "六";
            }
            else
            {
                return "日";
            }
            #endregion
        }
    }
}
