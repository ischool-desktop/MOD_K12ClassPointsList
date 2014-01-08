using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.IO;
using System.Xml;
using K12.Data.Configuration;
using K12.Data;
using Aspose.Words;
using System.Diagnostics;
using FISCA.DSAUtil;

namespace K12ClassPointsList.Modules
{
    public partial class ClassPointsList : BaseForm
    {
        //班級點名單設定內容
        /// <summary>
        /// 樣版
        /// </summary>
        string ClassPrint_Config_1 = "K12ClassPointsList.Modules.ClassPointsList.cs_1";

        BackgroundWorker BGW = new BackgroundWorker();

        List<string> config = new List<string>();

        int 學生多少個 = 100;
        int 日期多少天 = 60;

        public ClassPointsList()
        {
            InitializeComponent();
        }

        private void ClassPointsList_Load(object sender, EventArgs e)
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            dateTimeInput1.Value = DateTime.Today.AddDays(1);
            dateTimeInput2.Value = DateTime.Today.AddDays(7);

            GetDateTime_Click(null, null);
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (K12.Presentation.NLDPanels.Class.SelectedSource.Count == 0)
            {
                MsgBox.Show("請選擇班級!!");
                return;
            }

            if (BGW.IsBusy)
            {
                MsgBox.Show("忙碌中,稍後再試!!");
                return;
            }

            #region 日期設定

            if (dataGridViewX1.Rows.Count <= 0)
            {
                MsgBox.Show("列印點名單必須有日期!!");
                return;
            }

            DSXmlHelper dxXml = new DSXmlHelper("XmlData");

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                string 日期 = "" + row.Cells[0].Value;
                dxXml.AddElement(".", "item", 日期);
            }

            #endregion

            btnPrint.Enabled = false;
            BGW.RunWorkerAsync(dxXml.BaseElement);
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 取得樣版

            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(ClassPrint_Config_1);
            Aspose.Words.Document Template;

            if (ConfigurationInCadre.Template == null)
            {
                //如果範本為空,則建立一個預設範本
                Campus.Report.ReportConfiguration ConfigurationInCadre_1 = new Campus.Report.ReportConfiguration(ClassPrint_Config_1);
                ConfigurationInCadre_1.Template = new Campus.Report.ReportTemplate(Properties.Resources.點名表_套表列印_範本, Campus.Report.TemplateType.Word);
                Template = ConfigurationInCadre_1.Template.ToDocument();
            }
            else
            {
                //如果已有範本,則取得樣板
                Template = ConfigurationInCadre.Template.ToDocument();
            }

            SetTemplate(Template);

            #endregion

            //取得資料
            List<string> ClassIDList = K12.Presentation.NLDPanels.Class.SelectedSource;

            //每班一張,需要建立一個< 班級ID:List<學生ID> >


            //班級,座號,姓名,學號,性別
            List<StudRecord> StudList = new List<StudRecord>();
            StringBuilder sb = new StringBuilder();
            sb.Append("select class.id as class_id,class.display_order,class.class_name,class.grade_year,student.id as student_id,student.name,student.gender,student.seat_no,student.student_number,student.status,teacher.teacher_name ");
            sb.Append("from class join student on class.id=student.ref_class_id left join teacher on class.ref_teacher_id=teacher.id ");
            sb.Append(string.Format("where class.id in ('{0}') ", string.Join("','", ClassIDList)));
            sb.Append("and (student.status='1' or student.status='2') ");
            sb.Append("ORDER by class.grade_year,class.display_order,class.class_name,student.seat_no,student.name");

            DataTable dt = StatTool.Q.Select(sb.ToString());

            foreach (DataRow row in dt.Rows)
            {
                StudList.Add(new StudRecord(row));
            }

            Dictionary<string, List<StudRecord>> dic = new Dictionary<string, List<StudRecord>>();

            foreach (StudRecord each in StudList)
            {
                if (!dic.ContainsKey(each.班級名稱))
                {
                    dic.Add(each.班級名稱, new List<StudRecord>());
                }
                dic[each.班級名稱].Add(each);

            }

            #region 日期


            XmlElement day = (XmlElement)e.Argument;

            if (day == null)
            {
                MsgBox.Show("第一次使用報表請先進行[日期設定]");
                return;
            }
            else
            {
                config.Clear();
                foreach (XmlElement xml in day.SelectNodes("item"))
                {
                    config.Add(xml.InnerText);
                }
            }

            #endregion

            DataTable table = new DataTable();

            table.Columns.Add("學校名稱");
            table.Columns.Add("班級");
            table.Columns.Add("教師");
            table.Columns.Add("學年度");
            table.Columns.Add("學期");
            table.Columns.Add("列印日期");
            table.Columns.Add("上課開始");
            table.Columns.Add("上課結束");

            for (int x = 1; x <= 日期多少天; x++)
            {
                table.Columns.Add(string.Format("日期_{0}", x));
            }

            for (int x = 1; x <= 學生多少個; x++)
            {
                table.Columns.Add(string.Format("座號_{0}", x));
            }

            for (int x = 1; x <= 學生多少個; x++)
            {
                table.Columns.Add(string.Format("姓名_{0}", x));
            }

            for (int x = 1; x <= 學生多少個; x++)
            {
                table.Columns.Add(string.Format("學號_{0}", x));
            }

            for (int x = 1; x <= 學生多少個; x++)
            {
                table.Columns.Add(string.Format("性別_{0}", x));
            }

            //每個班級
            foreach (string className in dic.Keys)
            {
                DataRow row = table.NewRow();
                row["學校名稱"] = K12.Data.School.ChineseName;
                row["學年度"] = K12.Data.School.DefaultSchoolYear;
                row["學期"] = K12.Data.School.DefaultSemester;
                row["班級"] = className;
                row["列印日期"] = DateTime.Today.ToShortDateString();

                row["上課開始"] = config[0];
                row["上課結束"] = config[config.Count - 1];

                List<StudRecord> list = dic[className];
                row["教師"] = list[0].教師;

                for (int x = 1; x <= config.Count; x++)
                {
                    row[string.Format("日期_{0}", x)] = config[x - 1];
                }

                //每名學生
                int y = 1;
                foreach (StudRecord stud in list)
                {
                    if (y <= 學生多少個) //限制畫面到100名學生
                    {
                        row[string.Format("座號_{0}", y)] = stud.座號;
                        row[string.Format("姓名_{0}", y)] = stud.姓名;
                        row[string.Format("學號_{0}", y)] = stud.學號;
                        row[string.Format("性別_{0}", y)] = stud.性別;
                        y++;
                    }
                }

                table.Rows.Add(row);
            }


            Document PageOne = (Document)Template.Clone(true);
            PageOne.MailMerge.Execute(table);
            e.Result = PageOne;
        }

        private void SetColumn(DataTable table, int p, string p_2)
        {
            for (int x = 1; x <= p; x++)
            {
                table.Columns.Add(string.Format(p_2, x));
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnPrint.Enabled = true;

            if (e.Cancelled)
            {
                MsgBox.Show("作業已被中止!!");
            }
            else
            {
                if (e.Error == null)
                {
                    Document inResult = (Document)e.Result;

                    try
                    {
                        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                        SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                        SaveFileDialog1.FileName = "班級點名單(套表列印)";

                        if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            inResult.Save(SaveFileDialog1.FileName);
                            Process.Start(SaveFileDialog1.FileName);
                        }
                        else
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                            return;
                        }
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                        return;
                    }

                    this.Close();
                }
                else
                {
                    MsgBox.Show("列印資料發生錯誤\n" + e.Error.Message);
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            //取得設定檔
            Campus.Report.ReportConfiguration ConfigurationInCadre = new Campus.Report.ReportConfiguration(ClassPrint_Config_1);
            Campus.Report.TemplateSettingForm TemplateForm;
            //畫面內容(範本內容,預設樣式
            if (ConfigurationInCadre.Template != null)
            {
                TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(Properties.Resources.點名表_套表列印_範本, Campus.Report.TemplateType.Word));
            }
            else
            {
                ConfigurationInCadre.Template = new Campus.Report.ReportTemplate(Properties.Resources.點名表_套表列印_範本, Campus.Report.TemplateType.Word);
                TemplateForm = new Campus.Report.TemplateSettingForm(ConfigurationInCadre.Template, new Campus.Report.ReportTemplate(Properties.Resources.點名表_套表列印_範本, Campus.Report.TemplateType.Word));
            }

            //預設名稱
            TemplateForm.DefaultFileName = "班級點名單(套表列印範本)";

            //如果回傳為OK
            if (TemplateForm.ShowDialog() == DialogResult.OK)
            {
                SetTemplate(TemplateForm.Template.ToDocument());
                //設定後樣試,回傳
                ConfigurationInCadre.Template = TemplateForm.Template;
                //儲存
                ConfigurationInCadre.Save();
            }
        }

        private void SetTemplate(Document tp)
        {
            #region 使用者上傳的範本內學生數

            List<string> FieldNames = new List<string>(tp.MailMerge.GetFieldNames());
            int 使用者上傳的範本內學生數 = 1;
            foreach (string each in FieldNames)
            {
                if (each == string.Format("姓名_{0}", 使用者上傳的範本內學生數))
                {
                    使用者上傳的範本內學生數++;
                }
            }

            #endregion

            #region 使用者所選的班級學生最大數

            List<string> ClassIDList = K12.Presentation.NLDPanels.Class.SelectedSource;
            StringBuilder sb = new StringBuilder();
            sb.Append("select class.id,class.class_name,count(student.id) ");
            sb.Append("from student inner join class on student.ref_class_id=class.id ");
            sb.Append(string.Format("where class.id in ('{0}') and student.status in (1,2)", string.Join("','", ClassIDList)));
            sb.Append("group by class.id,class.class_name ");
            sb.Append("order by count(student.id) desc");

            DataTable dt = StatTool.Q.Select(sb.ToString());
            int 使用者所選的班級學生最大數 = 0;
            foreach (DataRow row in dt.Rows)
            {
                if (使用者所選的班級學生最大數 < int.Parse("" + row["count"]))
                {
                    使用者所選的班級學生最大數 = int.Parse("" + row["count"]);
                }
            }

            #endregion

            //比較人數
            if (使用者上傳的範本內學生數 < 使用者所選的班級學生最大數)
            {
                StringBuilder sb1 = new StringBuilder();
                sb1.AppendLine(string.Format("您目前所選的班級學生數最高為: {0} 名學生", 使用者所選的班級學生最大數));
                sb1.AppendLine(string.Format("大於上傳範本的最高學生可容納數: {0} 名學生", 使用者上傳的範本內學生數));
                sb1.AppendLine(string.Format("超過的 {0}名學生 將無法被列印出來!", 使用者所選的班級學生最大數 - 使用者上傳的範本內學生數));
                MsgBox.Show(sb1.ToString(), "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "另存新檔";
            sfd.FileName = "班級點名單_合併欄位總表.doc";
            sfd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Properties.Resources.合併欄位總表, 0, Properties.Resources.合併欄位總表.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void GetDateTime_Click(object sender, EventArgs e)
        {
            //建立日期清單
            TimeSpan ts = dateTimeInput2.Value - dateTimeInput1.Value;

            List<DateTime> TList = new List<DateTime>();

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                DateTime dt;
                DateTime.TryParse("" + row.Cells[0].Value, out dt);
                TList.Add(dt);
            }

            for (int x = 0; x <= ts.Days; x++)
            {
                DateTime dt = dateTimeInput1.Value.AddDays(x);
                //如果有排除假日,則排除Saturday and Sunday
                if (checkBoxX1.Checked)
                {
                    if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                    {
                        if (!TList.Contains(dt))
                        {
                            TList.Add(dt);
                        }
                    }
                }
                else
                {
                    if (!TList.Contains(dt))
                    {
                        TList.Add(dt);
                    }
                }
            }

            TList.Sort();
            //資料填入
            dataGridViewX1.Rows.Clear();
            foreach (DateTime dt in TList)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);
                row.Cells[0].Value = dt.ToShortDateString();
                row.Cells[1].Value = StatTool.CheckWeek(dt.DayOfWeek.ToString());
                dataGridViewX1.Rows.Add(row);
            }

        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            dataGridViewX1.Rows.Clear();
        }
    }
}
