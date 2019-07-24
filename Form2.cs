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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();

			comboBox1.Items.Add("LS成型機");
			comboBox1.Items.Add("NQD成型機");
			comboBox1.Items.Add("HS成型機");
			comboBox1.Items.Add("MS成型機");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string seikeiki = string.Format("成形機は「{0}」で合っていますか？",comboBox1.Text);
            DialogResult result = MessageBox.Show(seikeiki, "PGM成型機　選択", MessageBoxButtons.YesNo);

			if(result == DialogResult.Yes)
			{

	            //OpenFileDialogクラスのインスタンスを作成
	            OpenFileDialog ofd = new OpenFileDialog();

	            //はじめのファイル名を指定する
	            //はじめに「ファイル名」で表示される文字列を指定する
	            ofd.FileName = "PgmOut.csv";
	            //はじめに表示されるフォルダを指定する
	            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
	            ofd.InitialDirectory = @"C:\";
	            //[ファイルの種類]に表示される選択肢を指定する

	            //指定しないとすべてのファイルが表示される
				string seikeikiHeader = "";
				switch(comboBox1.SelectedIndex)
				{
				case 0:
					seikeikiHeader = "*LS*";
					break;
				case 1:
					seikeikiHeader = "*NQD*";
					break;
				case 2:
					seikeikiHeader = "*HS*";
					break;
				case 3:
					seikeikiHeader = "*MS*";
					break;
				default:
					break;
				}

	            ofd.Filter = "CSVファイル(*.csv)|" + seikeikiHeader + ".csv|すべてのファイル(*.*)|*.*";

	            //[ファイルの種類]ではじめに選択されるものを指定する
	            //2番目の「CSVファイル」が選択されているようにする
	            ofd.FilterIndex = 1;
	            //タイトルを設定する
	            ofd.Title = "開くファイルを選択してください";
	            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
	            ofd.RestoreDirectory = true;
	            //存在しないファイルの名前が指定されたとき警告を表示する
	            //デフォルトでTrueなので指定する必要はない
	            ofd.CheckFileExists = true;
	            //存在しないパスが指定されたとき警告を表示する
	            //デフォルトでTrueなので指定する必要はない
	            ofd.CheckPathExists = true;

	            //ダイアログを表示する
	            if (ofd.ShowDialog() == DialogResult.OK)
	            {
	                //OKボタンがクリックされたとき、選択されたファイル名を表示する
	                int sel = comboBox1.SelectedIndex;
	                if(sel == 0 || sel == 1)//LS NQD
	                {
		                Form1 form1 = new Form1();
		                form1.SetSelectedFile(ofd.FileName);
		                form1.SetSelectedHeader(seikeikiHeader);
		                form1.Show();
		                this.Hide();
		            }
	                else if(sel == 2 || sel == 3)//HS MS
	                {
		                Form4 form4 = new Form4();
		                form4.SetSelectedFile(ofd.FileName);
		                form4.SetSelectedHeader(seikeikiHeader);
		                form4.Show();
		                this.Hide();
					}

	            }

			}

        }
    }
}
