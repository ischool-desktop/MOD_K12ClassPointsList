using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace K12ClassPointsList.Modules
{
    class StudRecord
    {
        public string 班級ID { get; set; }
        public string 班級名稱 { get; set; }
        public string 年級 { get; set; }
        public string 班級序號 { get; set; }

        public string 學生ID { get; set; }
        public string 姓名 { get; set; }
        public string 座號 { get; set; }
        public string 學號 { get; set; }


        private string _性別 { get; set; }

        public string 性別
        {
            get
            {
                if (_性別 == "1")
                    return "男";
                else if (_性別 == "0")
                    return "女";
                else
                    return "";
            }
            set
            {
                _性別 = value;
            }
        }

        public string 狀態 { get; set; }

        public string 教師 { get; set; }

        public StudRecord(DataRow row)
        {

            班級ID = "" + row["class_id"];
            班級名稱 = "" + row["class_name"];
            年級 = "" + row["grade_year"];
            班級序號 = "" + row["display_order"];

            學生ID = "" + row["student_id"];
            姓名 = "" + row["name"];
            座號 = "" + row["seat_no"];
            學號 = "" + row["student_number"];
            性別 = "" + row["gender"];
            狀態 = "" + row["status"];
            教師 = "" + row["teacher_name"];
        }
    }
}
