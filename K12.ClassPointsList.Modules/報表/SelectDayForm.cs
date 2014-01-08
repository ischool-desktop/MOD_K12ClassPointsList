using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Data.Configuration;
using System.Xml;
using FISCA.DSAUtil;

namespace K12ClassPointsList.Modules
{
    public partial class SelectDayForm : BaseForm
    {

        string _config { get; set; }

        List<string> list { get; set; }

        public SelectDayForm()
        {
            InitializeComponent();
        }

        public SelectDayForm(string config)
        {
            InitializeComponent();

            _config = config;
            dateTimeInput1.Value = DateTime.Today;
            dateTimeInput2.Value = DateTime.Today;
        }

        private void SelectDayForm_Load(object sender, EventArgs e)
        {
            list = new List<string>();
            //取得使用者所設定之日期清單
            K12.Data.Configuration.ConfigData cd = School.Configuration[_config];

            XmlElement preferenceData = cd.GetXml("日期設定", null);

            if (preferenceData != null)
            {
                foreach (XmlElement period in preferenceData.SelectNodes("item"))
                {
                    string name = period.InnerText;
                    list.Add(name);
                }
            }

            foreach (string each in list)
            {
                DateTime dt;

                DateTime.TryParse(each, out dt);

                if (dt != null)
                {
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = dt.ToShortDateString();
                    row.Cells[1].Value = StatTool.CheckWeek(dt.DayOfWeek.ToString());
                    dataGridViewX1.Rows.Add(row);
                }
            }
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



        private void btnSave_Click(object sender, EventArgs e)
        {
            DSXmlHelper dxXml = new DSXmlHelper("XmlData");

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                string 日期 = "" + row.Cells[0].Value;
                dxXml.AddElement(".", "item", 日期);
            }

            K12.Data.Configuration.ConfigData cd = School.Configuration[_config];
            cd["日期設定"] = dxXml.BaseElement.OuterXml;
            cd.Save();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            dataGridViewX1.Rows.Clear();
        }
    }
}
