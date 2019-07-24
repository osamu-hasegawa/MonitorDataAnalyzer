using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MonitorDataAnalyzer
{
    public partial class FormCalender : Form
    {
        Form1 form1 = null;
        Form4 form4 = null;
		int formType = 0;
        
        public FormCalender()
        {
            InitializeComponent();

            label18.Text = "";
            label19.Text = "";
            label20.Text = "";
            label22.Text = "";

			label36.Text = "";
			label37.Text = "";
			label38.Text = "";
			label39.Text = "";
			label40.Text = "";
			label41.Text = "";
			label42.Text = "";
			label43.Text = "";
			label44.Text = "";
			label45.Text = "";
			label46.Text = "";
			label47.Text = "";
			label48.Text = "";

			label50.Text = "0";
			label51.Text = "0";
			label52.Text = "0";
			label53.Text = "0";

            monthCalendar1.Font = new Font("MS UI Gothic", 20, FontStyle.Bold);
        }

        public void SetParentForm(Form form, int type)
        {
			formType = type;
			if(formType == 1)//LS,NQD
			{
	            form1 = (Form1)form;
	        }
	        else if(formType == 2)//HS,MS
	        {
	            form4 = (Form4)form;
	        }
				
        }
        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
		    DateTime start = e.Start;
		    DateTime end = e.End;

			if(start == end)//1回目は現在時刻となる為、修正
			{
				string ss = "00:00:00";
				start = DateTime.Parse(ss);
				string se = "23:59:59";
				end = DateTime.Parse(se);
			}

            string stDate = start.ToString("yyyy/MM/dd HH:mm:ss");
            string enDate = end.ToString("yyyy/MM/dd HH:mm:ss");

		    DateTime selectStaDate = e.Start;
		    DateTime selectEndDate = e.End;
            label49.Text = selectStaDate.ToString("yyyy/MM/dd HH:mm:ss") + "～" + "\r\n" + selectEndDate.ToString("yyyy/MM/dd HH:mm:ss");

            label18.Text = "0";
            label19.Text = "0";
            label20.Text = "0";
            label22.Text = "0";

			label36.Text = "0";
			label37.Text = "0";
			label38.Text = "0";
			label39.Text = "0";
			label40.Text = "0";
			label41.Text = "0";
			label42.Text = "0";
			label43.Text = "0";
			label44.Text = "0";
			label45.Text = "0";
			label46.Text = "0";
			label47.Text = "0";
			label48.Text = "0";

			label50.Text = "0";
			label51.Text = "0";
			label52.Text = "0";
			label53.Text = "0";

			label54.Text = "";

            string okngStr = "";
            string ngcauseStr = "";
            string nikuInfo = "";
            string sleeveInfo = "";
            double kadou = 0;

            if(formType == 1)//LS,NQD
			{
	            form1.SetDateTimeToAnalize(stDate, enDate, ref okngStr, ref ngcauseStr, ref nikuInfo, ref sleeveInfo, ref kadou);
			}
			else if(formType == 2)//HS,MS
			{
	            form4.SetDateTimeToAnalize(stDate, enDate, ref okngStr, ref ngcauseStr, ref nikuInfo, ref sleeveInfo, ref kadou);
	        }

            string[] line = okngStr.Split(',');

            label18.Text = line[0] + "個";
	        label19.Text = line[1] + "個";
			label18.ForeColor = Color.Green;
			label19.ForeColor = Color.Red;
            int okCount = int.Parse(line[0]);
            int ngCount = int.Parse(line[1]);
			if(okCount != 0 && ngCount != 0)
			{
				double oknum = (double)okCount;
				double ngnum = (double)ngCount;
                double total = (double)(oknum + ngnum);
                double budomari = (double)(oknum / total * 100);
	            label20.Text = budomari.ToString("F1") + "％";
				label22.Text = total.ToString() + "個";
	        }

			Label [] item = new Label [] {label23, label24, label25, label26, label27, label28, label29, label30, label31, label32, label33, label34, label35};
			Label [] valu = new Label [] {label36, label37, label38, label39, label40, label41, label42, label43, label44, label45, label46, label47, label48};

			for(int i = 0; i < item.Length; i++)
			{
				item[i].ForeColor = Color.Black;
				valu[i].ForeColor = Color.Black;
			}

			if(ngCount > 0)
			{
                string[] ngline = ngcauseStr.Split(',');

				for(int i = 0; i < item.Length; i++)
				{
					valu[i].Text = ngline[i] + "個";
					if(int.Parse(ngline[i]) > 0)
					{
						item[i].ForeColor = Color.Red;
						valu[i].ForeColor = Color.Red;
					}
				}
		    }

			if(nikuInfo != "")
			{
                string[] nikuline = nikuInfo.Split(',');

				label51.Text = nikuline[0];//平均
				label52.Text = nikuline[1];//最大
				label50.Text = nikuline[2];//最小
				label53.Text = nikuline[3];//標準偏差
			}


			Label [] label = new Label[]{label12, label11, label10, label9, label16, label15, label14, label13};
			for(int i = 0; i < label.Length; i++)
			{
				label[i].Text = "";
			}

			if(sleeveInfo != "")
			{
                string[] sleeveline = sleeveInfo.Split(',');
                int j = 0;
				for(int i = 0; i < sleeveline.Length; i+=2)
				{
					label[j].Text = "SL-" + sleeveline[i] + " : " + sleeveline[i + 1];
					j++;
				}
			}
			
			if(!Double.IsNaN(kadou))
			{
				label54.Text = "稼働率(使用数/全数)：" + kadou.ToString("F3");
			}

        }

    }
}
