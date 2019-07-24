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
    public partial class Form3 : Form
    {
		string displayString = "";
        Label[] labelTable;

        public Form3()
        {
            InitializeComponent();
        }

		public void SetDetailData(string detailData)
		{
			displayString = detailData;
			string[] fields = displayString.Split(',');

            labelTable = new Label[] {this.label1, this.label2, this.label3, this.label4, this.label5, this.label6, this.label7, this.label8, this.label9, this.label10,
                                        this.label11, this.label12, this.label13, this.label14, this.label15, this.label16, this.label17, this.label18, this.label19, this.label20,
                                        this.label21, this.label22, this.label23, this.label24, this.label25, this.label26, this.label27, /*this.label28,*/ this.label29, this.label30};

			for(int i = 0; i < labelTable.Length; i++)
			{
				labelTable[i].Text = fields[i];
                if(labelTable[24].Text == "OK")
                {
					labelTable[24].ForeColor = Color.Green;
				}
                else if(labelTable[24].Text == "NG")
                {
					labelTable[24].ForeColor = Color.Red;
				}
			}


        }

    }
}
