using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Microsoft.VisualBasic.FileIO;

namespace MonitorDataAnalyzer
{
    public partial class Form4 : Form
    {

		public struct SleeveList
		{
			public string sleeveNumber;
			public int shotCount;
		}

		string selectedFile = "";

		string Date = "";
		string Time = "";
		string Seikeiki = "";
		string SeihinName = "";

		string CsvFileName = "";

		string DateAndTime = "";
		string Hinshumei = "";
		string TotalCount = "";
		string Sampling = "";
		string ProcessTime = "";

		string Seikei1No = "";
		string Seikei1Count = "";
		string Seikei1Tkeisu = "";
		string Seikei1Zhosei = "";
		string Seikei1ct1up = "";
		string Seikei1ct1dn = "";
		string Seikei1Z = "";
		string Seikei1cc2cc1 = "";
		string Seikei1cp = "";
		string Seikei1step1 = "";
		string Seikei1step2 = "";
		string Seikei1step3 = "";
		string Seikei1step4 = "";
		string Seikei1step5 = "";
		string Seikei1step6 = "";
		string Seikei1step7 = "";
		string Seikei1step8 = "";
		string Seikei1step9 = "";
		string Seikei1step10 = "";
		string Seikei1step11 = "";
		string Seikei1step12 = "";
		string Seikei1Total = "";

		string Seikei2No = "";
		string Seikei2Count = "";
		string Seikei2Tkeisu = "";
		string Seikei2Zhosei = "";
		string Seikei2ct1up = "";
		string Seikei2ct1dn = "";
		string Seikei2Z = "";
		string Seikei2cc2cc1 = "";
		string Seikei2cp = "";
		string Seikei2step1 = "";
		string Seikei2step2 = "";
		string Seikei2step3 = "";
		string Seikei2step4 = "";
		string Seikei2step5 = "";
		string Seikei2step6 = "";
		string Seikei2step7 = "";
		string Seikei2step8 = "";
		string Seikei2step9 = "";
		string Seikei2step10 = "";
		string Seikei2step11 = "";
		string Seikei2step12 = "";
		string Seikei2Total = "";

		string RoundCount = "";
		string Error = "";

		string NikuatsuUpper = "";
		string NikuatsuValue = "";
		string NikuatsuLower = "";
		string Result = "";
		string NgCause = "";
		string Housharitsu = "";
		string Operator = "";
		string ResistSleeve = "";
		string ShotCount = "";
		string SeikeiCount = "";
	
		string seikeikiHeader = "";
		string dstFile = "";

		FileStream fp;
		StreamReader sr = null;

		FormCalender formCalender = null;

        public Form4()
        {
            InitializeComponent();

            // ListViewコントロールのプロパティを設定
            listView1.FullRowSelect = true;
            listView1.GridLines = true;
            listView1.Sorting = SortOrder.Ascending;
            listView1.View = View.Details;
            listView1.HideSelection = false;

            // 列（コラム）ヘッダの作成
            ColumnHeader columnDate = new ColumnHeader();
            ColumnHeader columnTime = new ColumnHeader();
            ColumnHeader columnSeikeiki = new ColumnHeader();
            ColumnHeader columnSeihinName = new ColumnHeader();
            ColumnHeader columnCsvFileName = new ColumnHeader();

            ColumnHeader columnDateTime = new ColumnHeader();
            ColumnHeader columnHinshumei = new ColumnHeader();
            ColumnHeader columnTotalCount = new ColumnHeader();
            ColumnHeader columnSampling = new ColumnHeader();
            ColumnHeader columnProcessTime = new ColumnHeader();

            ColumnHeader columnSeikei1No = new ColumnHeader();
            ColumnHeader columnSeikei1Count = new ColumnHeader();
            ColumnHeader columnSeikei1Tkeisu = new ColumnHeader();
            ColumnHeader columnSeikei1Zhosei = new ColumnHeader();
            ColumnHeader columnSeikei1ct1up = new ColumnHeader();
            ColumnHeader columnSeikei1ct1dn = new ColumnHeader();
            ColumnHeader columnSeikei1Z = new ColumnHeader();
            ColumnHeader columnSeikei1cc2cc1 = new ColumnHeader();
            ColumnHeader columnSeikei1cp = new ColumnHeader();
            ColumnHeader columnSeikei1step1 = new ColumnHeader();
            ColumnHeader columnSeikei1step2 = new ColumnHeader();
            ColumnHeader columnSeikei1step3 = new ColumnHeader();
            ColumnHeader columnSeikei1step4 = new ColumnHeader();
            ColumnHeader columnSeikei1step5 = new ColumnHeader();
            ColumnHeader columnSeikei1step6 = new ColumnHeader();
            ColumnHeader columnSeikei1step7 = new ColumnHeader();
            ColumnHeader columnSeikei1step8 = new ColumnHeader();
            ColumnHeader columnSeikei1step9 = new ColumnHeader();
            ColumnHeader columnSeikei1step10 = new ColumnHeader();
            ColumnHeader columnSeikei1step11 = new ColumnHeader();
            ColumnHeader columnSeikei1step12 = new ColumnHeader();
            ColumnHeader columnSeikei1Total = new ColumnHeader();

            ColumnHeader columnSeikei2No = new ColumnHeader();
            ColumnHeader columnSeikei2Count = new ColumnHeader();
            ColumnHeader columnSeikei2Tkeisu = new ColumnHeader();
            ColumnHeader columnSeikei2Zhosei = new ColumnHeader();
            ColumnHeader columnSeikei2ct1up = new ColumnHeader();
            ColumnHeader columnSeikei2ct1dn = new ColumnHeader();
            ColumnHeader columnSeikei2Z = new ColumnHeader();
            ColumnHeader columnSeikei2cc2cc1 = new ColumnHeader();
            ColumnHeader columnSeikei2cp = new ColumnHeader();
            ColumnHeader columnSeikei2step1 = new ColumnHeader();
            ColumnHeader columnSeikei2step2 = new ColumnHeader();
            ColumnHeader columnSeikei2step3 = new ColumnHeader();
            ColumnHeader columnSeikei2step4 = new ColumnHeader();
            ColumnHeader columnSeikei2step5 = new ColumnHeader();
            ColumnHeader columnSeikei2step6 = new ColumnHeader();
            ColumnHeader columnSeikei2step7 = new ColumnHeader();
            ColumnHeader columnSeikei2step8 = new ColumnHeader();
            ColumnHeader columnSeikei2step9 = new ColumnHeader();
            ColumnHeader columnSeikei2step10 = new ColumnHeader();
            ColumnHeader columnSeikei2step11 = new ColumnHeader();
            ColumnHeader columnSeikei2step12 = new ColumnHeader();
            ColumnHeader columnSeikei2Total = new ColumnHeader();

            ColumnHeader columnRoundCount = new ColumnHeader();
            ColumnHeader columnError = new ColumnHeader();

            ColumnHeader columnNikuatsuUpper = new ColumnHeader();
            ColumnHeader columnNikuatsuValue = new ColumnHeader();
            ColumnHeader columnNikuatsuLower = new ColumnHeader();
            ColumnHeader columnResult = new ColumnHeader();
            ColumnHeader columnNgCause = new ColumnHeader();
            ColumnHeader columnHousharitsu = new ColumnHeader();
            ColumnHeader columnOperator = new ColumnHeader();
            ColumnHeader columnResistSleeve = new ColumnHeader();
            ColumnHeader columnShotCount = new ColumnHeader();
            ColumnHeader columnSeikeiCount = new ColumnHeader();

			columnDate.Text = "日付";
			columnTime.Text = "時間";
			columnSeikeiki.Text = "成型機";
			columnSeihinName.Text = "製品名";
			columnCsvFileName.Text = "ﾛｸﾞﾌｧｲﾙ名";
			columnDateTime.Text = "日時";
			columnHinshumei.Text = "品種名";
			columnTotalCount.Text = "総投入数";
			columnSampling.Text = "サンプリングデータ数";
			columnProcessTime.Text = "処理時間";

			columnSeikei1No.Text = "成型1金型番号";
			columnSeikei1Count.Text = "成型1金型投入回数";
			columnSeikei1Tkeisu.Text = "成型1T係数";
			columnSeikei1Zhosei.Text = "成型1Z補正";
			columnSeikei1ct1up.Text = "成型1ct1(上)";
			columnSeikei1ct1dn.Text = "成型1ct1(下)";
			columnSeikei1Z.Text = "成型1Z";
			columnSeikei1cc2cc1.Text = "成型1cc2-cc1";
			columnSeikei1cp.Text = "成型1cp";
			columnSeikei1step1.Text = "成型1Step1";
			columnSeikei1step2.Text = "成型1Step2";
			columnSeikei1step3.Text = "成型1Step3";
			columnSeikei1step4.Text = "成型1Step4";
			columnSeikei1step5.Text = "成型1Step5";
			columnSeikei1step6.Text = "成型1Step6";
			columnSeikei1step7.Text = "成型1Step7";
			columnSeikei1step8.Text = "成型1Step8";
			columnSeikei1step9.Text = "成型1Step9";
			columnSeikei1step10.Text = "成型1Step10";
			columnSeikei1step11.Text = "成型1Step11";
			columnSeikei1step12.Text = "成型1Step12";
			columnSeikei1Total.Text = "成型1Total";

			columnSeikei2No.Text = "成型2金型番号";
			columnSeikei2Count.Text = "成型2金型投入回数";
			columnSeikei2Tkeisu.Text = "成型2T係数";
			columnSeikei2Zhosei.Text = "成型2Z補正";
			columnSeikei2ct1up.Text = "成型2ct1(上)";
			columnSeikei2ct1dn.Text = "成型2ct1(下)";
			columnSeikei2Z.Text = "成型2Z";
			columnSeikei2cc2cc1.Text = "成型2cc2-cc1";
			columnSeikei2cp.Text = "成型2cp";
			columnSeikei2step1.Text = "成型2Step1";
			columnSeikei2step2.Text = "成型2Step2";
			columnSeikei2step3.Text = "成型2Step3";
			columnSeikei2step4.Text = "成型2Step4";
			columnSeikei2step5.Text = "成型2Step5";
			columnSeikei2step6.Text = "成型2Step6";
			columnSeikei2step7.Text = "成型2Step7";
			columnSeikei2step8.Text = "成型2Step8";
			columnSeikei2step9.Text = "成型2Step9";
			columnSeikei2step10.Text = "成型2Step10";
			columnSeikei2step11.Text = "成型2Step11";
			columnSeikei2step12.Text = "成型2Step12";
			columnSeikei2Total.Text = "成型2Total";

			columnRoundCount.Text = "周回カウント";
			columnError.Text = "異常";

			columnNikuatsuUpper.Text = "肉厚上限";
			columnNikuatsuValue.Text = "肉厚測定値";
			columnNikuatsuLower.Text = "肉厚下限";
			columnResult.Text = "合否";
			columnNgCause.Text = "不良原因";
			columnHousharitsu.Text = "放射率";
			columnOperator.Text = "作業者";
			columnResistSleeve.Text = "ｽﾘｰﾌﾞNo";
			columnShotCount.Text = "ｼｮｯﾄ数";
			columnSeikeiCount.Text = "成型数";

            ColumnHeader[] colHeaderRegValue =
			{
				columnDate, columnTime, columnSeikeiki, columnSeihinName, columnCsvFileName, columnDateTime, columnHinshumei, columnTotalCount, columnSampling, columnProcessTime, 
				columnSeikei1No, columnSeikei1Count, columnSeikei1Tkeisu, columnSeikei1Zhosei, columnSeikei1ct1up, columnSeikei1ct1dn, columnSeikei1Z, columnSeikei1cc2cc1, columnSeikei1cp, 
				columnSeikei1step1, columnSeikei1step2, columnSeikei1step3, columnSeikei1step4, columnSeikei1step5, columnSeikei1step6, columnSeikei1step7, columnSeikei1step8, columnSeikei1step9, 
				columnSeikei1step10, columnSeikei1step11, columnSeikei1step12, columnSeikei1Total, columnSeikei2No, columnSeikei2Count, columnSeikei2Tkeisu, columnSeikei2Zhosei, columnSeikei2ct1up, 
				columnSeikei2ct1dn, columnSeikei2Z, columnSeikei2cc2cc1, columnSeikei2cp, columnSeikei2step1, columnSeikei2step2, columnSeikei2step3, columnSeikei2step4, columnSeikei2step5, 
				columnSeikei2step6, columnSeikei2step7, columnSeikei2step8, columnSeikei2step9, columnSeikei2step10, columnSeikei2step11, columnSeikei2step12, columnSeikei2Total, columnRoundCount, columnError,
				columnNikuatsuUpper, columnNikuatsuValue, columnNikuatsuLower, columnResult, columnNgCause, columnHousharitsu, columnOperator, columnResistSleeve, columnShotCount, columnSeikeiCount
			};

			listView1.Columns.AddRange(colHeaderRegValue);

			//ヘッダの幅を自動調節
			listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

			//ポイントで選択できるようにする
			listView1.HideSelection = false;

			//全体の色設定
            listView1.Sorting = SortOrder.None;
            listView1.ForeColor = Color.Black;//初期の色
            listView1.BackColor = Color.White;//背景色

			System.Reflection.Assembly asm = System.Reflection.Assembly.GetExecutingAssembly();
			System.Version ver = asm.GetName().Version;
            this.Text += "  Ver:" + ver;
        }

        private void Form4_FormClosed(object sender, FormClosedEventArgs e)
        {
			try
			{
				File.Delete(@dstFile);//一時ファイルは削除
			}
			catch(System.IO.IOException ex)
			{
				string errorLog = "他のアプリがtmpファイルを開いている可能性があります";
				System.Console.WriteLine(errorLog);
				System.Console.WriteLine(ex.Message);
				LogFileOut(errorLog);
				return;
			}

            string errorStr = "正常に終了しました。";
            LogFileOut(errorStr);

            Application.Exit();
        }

        public void SetSelectedFile(string file)
		{
			selectedFile = file;
			int pos = selectedFile.LastIndexOf("\\");
			string openFile = selectedFile.Substring(pos + 1);

			string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
            dstFile = stCurrentDir + "\\" + openFile + ".tmp";

			//tmpファイルがあれば、全て削除しておく
            System.IO.DirectoryInfo tmp_di = new System.IO.DirectoryInfo(stCurrentDir);
            IEnumerable<System.IO.FileInfo> files_tmp = tmp_di.EnumerateFiles("*.tmp", System.IO.SearchOption.AllDirectories);

			try
			{
				foreach(System.IO.FileInfo f in files_tmp)
				{
					string deleteFile = stCurrentDir + "\\" + f.Name;
					File.SetAttributes(deleteFile, FileAttributes.Normal);
					File.Delete(deleteFile);
				}
			}
			catch(System.IO.IOException ex)
			{
				string errorStr = "他のアプリがtmpファイルを開いている可能性があります";
				System.Console.WriteLine(errorStr);
				System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
				return;
			}

			//ファイルを開く
			try
			{
				//WORK用一時ファイルにコピーし、以降は一時ファイルを扱う。
				File.Copy(@selectedFile, @dstFile);
				fp = new FileStream(dstFile, FileMode.Open, FileAccess.Read);
				sr = new StreamReader(fp, System.Text.Encoding.GetEncoding("Shift_JIS"));
			}
			catch(System.IO.IOException ex)
			{
				string errorStr = "ファイルコピー失敗またはファイルが開けなかった可能性があります";
				System.Console.WriteLine(errorStr);
				System.Console.WriteLine(ex.Message);

				if(sr != null)
				{
					sr.Close();
				}
				File.Delete(@dstFile);//一時ファイルは削除
				LogFileOut(errorStr);
				return;
			}
			if (sr != null)
			{
				sr.Close();
			}

        }

        public void SetSelectedHeader(string header)
        {
            seikeikiHeader = header;
        }

        public void LogFileOut(string logMessage)
        {
            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
            string path = stCurrentDir + "\\MonitorDataAnalyzer.log";

            using (var sw = new System.IO.StreamWriter(path, true, System.Text.Encoding.Default))
            {
                sw.WriteLine($"{DateTime.Now.ToLongDateString()} {DateTime.Now.ToLongTimeString()}");
                sw.WriteLine($"  {logMessage}");
                sw.WriteLine("--------------------------------------------------------------");
            }
        }

        private void Form4_Shown(object sender, EventArgs e)
        {
			//ListViewを一度クリアする
            listView1.Items.Clear();

            string[] item1 = {
			Date, Time, Seikeiki, SeihinName, CsvFileName, DateAndTime, Hinshumei, TotalCount, Sampling, ProcessTime, 
			Seikei1No, Seikei1Count, Seikei1Tkeisu, Seikei1Zhosei, Seikei1ct1up, Seikei1ct1dn, Seikei1Z, Seikei1cc2cc1, Seikei1cp, Seikei1step1, Seikei1step2, 
			Seikei1step3, Seikei1step4, Seikei1step5, Seikei1step6, Seikei1step7, Seikei1step8, Seikei1step9, Seikei1step10, Seikei1step11, Seikei1step12, Seikei1Total, 
			Seikei2No, Seikei2Count, Seikei2Tkeisu, Seikei2Zhosei, Seikei2ct1up, Seikei2ct1dn, Seikei2Z, Seikei2cc2cc1, Seikei2cp, Seikei2step1, Seikei2step2, Seikei2step3, 
			Seikei2step4, Seikei2step5, Seikei2step6, Seikei2step7, Seikei2step8, Seikei2step9, Seikei2step10, Seikei2step11, Seikei2step12, Seikei2Total, RoundCount, Error, 
			NikuatsuUpper, NikuatsuValue, NikuatsuLower, Result, NgCause, Housharitsu, Operator, ResistSleeve, ShotCount, SeikeiCount
			};

            //CSV読込。高速化を狙い、最大行数を取得後、for分でループする
            var readToEnd = File.ReadAllLines(@dstFile, Encoding.GetEncoding("Shift_JIS"));
            int lines = readToEnd.Length;

            for (int i = 0; i < lines; i++)
            {
				if(i == 0)//ヘッダ部はスキップ
				{
					continue;
				}
                //１行のstringをstream化してTextFieldParserで処理する
                using (Stream stream = new MemoryStream(Encoding.Default.GetBytes(readToEnd[i])))
                {
                    using (TextFieldParser parser = new TextFieldParser(stream, Encoding.GetEncoding("Shift_JIS")))
                    {
                        parser.TextFieldType = FieldType.Delimited;
                        parser.Delimiters = new[] { "," };
                        parser.HasFieldsEnclosedInQuotes = true;
                        parser.TrimWhiteSpace = false;
                        string[] fields = parser.ReadFields();

                        for (int j = 0; j < fields.Length; j++)
                        {
							item1[j] = fields[j];
                        }
                    }

                }

                listView1.Items.Insert(0, new ListViewItem(item1));//先頭に追加
                listView1.Font = new System.Drawing.Font("Times New Roman", 10, System.Drawing.FontStyle.Regular);

                if (item1[59] == "OK")
                {
                    listView1.Items[0].BackColor = Color.Green;
                }
                else if (item1[59] == "NG")
                {
                    listView1.Items[0].BackColor = Color.Red;
                }
                else
                {
	                //日勤時間帯
	                string strNoonSta = "08:00:00";
	                string strNoonEnd = "17:00:00";
	                DateTime noonSta = DateTime.Parse(strNoonSta);
	                DateTime noonEnd = DateTime.Parse(strNoonEnd);

	                //夕勤時間帯
	                string strSunsetSta = "17:00:00";
	                string strSunsetEnd = "23:59:59";
	                DateTime sunsetSta = DateTime.Parse(strSunsetSta);
	                DateTime sunsetEnd = DateTime.Parse(strSunsetEnd);

	                //夜勤時間帯
	                string strNightSta = "00:00:01";
	                string strNightEnd = "08:00:00";
	                DateTime nightSta = DateTime.Parse(strNightSta);
	                DateTime nightEnd = DateTime.Parse(strNightEnd);

	                DateTime dt3 = DateTime.Parse(item1[1]);//タイムスタンプ(文字列)→DateTimeに変換

	                if (noonSta <= dt3 && dt3 < noonEnd)
	                {
	                    listView1.Items[0].ForeColor = Color.Blue;//青
	                }
	                if (sunsetSta <= dt3 && dt3 < sunsetEnd)
	                {
	                    listView1.Items[0].ForeColor = Color.Green;//緑
	                }
	                if (nightSta <= dt3 && dt3 < nightEnd)
	                {
	                    listView1.Items[0].ForeColor = Color.Red;//赤
	                }
	            }
            }

            //ヘッダの幅を自動調節
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private void button1_Click(object sender, EventArgs e)
        {
			if(formCalender == null || formCalender.IsDisposed)
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
	                //                Console.WriteLine(ofd.FileName);
	                this.SetSelectedFile(ofd.FileName);
	                this.Form4_Shown(sender, e);

	            }
			}
        }

        private void button2_Click(object sender, EventArgs e)
        {
			if(formCalender == null || formCalender.IsDisposed)
			{
				formCalender = new FormCalender();
				formCalender.Show();
				formCalender.SetParentForm(this, 2);
	        }
        }
        
        public void SetDateTimeToAnalize(string start, string end, ref string okngStr, ref string ngcauseStr, ref string nikuInfo, ref string sleeveInfo, ref double kadou)
        {
            DateTime dtsta = DateTime.Parse(start);
            DateTime dtend = DateTime.Parse(end);

	        int okCount = 0;
	        int ngCount = 0;

			int allSleeve = 0;
			int workSleeve = 0;
			int nikuatsuNG = 0;//肉厚不良
			int kizuNG = 0;//キズ
			int butsuNG = 0;//ブツ
			int yakeNG = 0;//ヤケ
			int hibicrackNG = 0;//ヒビ/クラック
			int gasukizuNG = 0;//ガスキズ
			int houshakizuNG = 0;//放射キズ
			int giratsukikumoriNG = 0;//ギラツキ/クモリ
			int hennikumendareNG = 0;//偏肉/面ダレ
			int yogoreNG = 0;//汚れ
			int hokoriNG = 0;//ほこり
			int keijoseidoNG = 0;//形状精度
			int etcNG = 0;//その他

            double nikuData = 0;
            double maxNiku = 0;
            double minNiku = 100;
            double sumNiku = 0;
            double sumV_Niku = 0;
			int nikuCount = 0;
			var nikuList = new List<double>();
			List<SleeveList> list = new List<SleeveList>();

            for (int i = 0; i < listView1.Items.Count; i++)
            {
                int colmax = listView1.Items[i].SubItems.Count;
                string listDate = "";
                string listTime = "";
                string listDT = "";
                string result = "";
                string ngcause = "";
                string shotCount = "";
                string seikeiCount = "";
				string nikuatsuSokutei = "";
				string sleeveNumber = "";
                for (int j = 0; j < colmax; j++)
                {           
                    string listitem = listView1.Items[i].SubItems[j].Text;
                    if (j == 0)//日付
                    {
                        listDate = listitem;
                    }
                    else if(j == 1)//時間
                    {
                        listTime = listitem;
                    }
                    else if(j == 57)//肉厚測定値
					{
						nikuatsuSokutei = listitem;
					}
                    else if(j == 59)//合否
					{
						result = listitem;
					}
                    else if(j == 60)//不良原因
					{
						ngcause = listitem;
					}
                    else if(j == 63)//スリーブ番号
					{
						sleeveNumber = listitem;
					}
                    else if(j == 64)//ショット数
					{
						shotCount = listitem;
					}
                    else if(j == 65)//成型数
					{
						seikeiCount = listitem;
					}

                }

                listDT = listDate + " " + listTime;
                DateTime dttarget = DateTime.Parse(listDT);
                if(dtsta < dttarget && dttarget < dtend)
                {
					allSleeve++;
					if(result == "OK")
					{
	                    okCount++;
	                }
					if(result == "NG")
					{
	                    ngCount++;

						if(ngcause == "肉厚不良")
						{
							nikuatsuNG++;//肉厚不良
						}
						else if(ngcause == "キズ")
						{
							kizuNG++;//キズ
						}
						else if(ngcause == "ブツ")
						{
							butsuNG++;//ブツ
						}
						else if(ngcause == "ヤケ")
						{
							yakeNG++;//ヤケ
						}
						else if(ngcause == "ヒビ/クラック")
						{
							hibicrackNG++;//ヒビ/クラック
						}
						else if(ngcause == "ガスキズ")
						{
							gasukizuNG++;//ガスキズ
						}
						else if(ngcause == "放射キズ")
						{
							houshakizuNG++;//放射キズ
						}
						else if(ngcause == "ギラツキ/クモリ")
						{
							giratsukikumoriNG++;//ギラツキ/クモリ
						}
						else if(ngcause == "偏肉/面ダレ")
						{
							hennikumendareNG++;//偏肉/面ダレ
						}
						else if(ngcause == "汚れ")
						{
							yogoreNG++;//汚れ
						}
						else if(ngcause == "ほこり")
						{
							hokoriNG++;//ほこり
						}
						else if(ngcause == "形状精度")
						{
							keijoseidoNG++;//形状精度
						}
						else if(ngcause == "その他")
						{
							etcNG++;//その他
						}
	                }
	                
	                if(nikuatsuSokutei != "-" && nikuatsuSokutei != "")
					{
		                nikuData = double.Parse(nikuatsuSokutei.ToString());
		                sumNiku += nikuData;
		                sumV_Niku += (nikuData * nikuData);
		                nikuList.Add(nikuData);
		             	nikuCount++;
		                
		                if(nikuData > maxNiku)
		                {
							maxNiku = nikuData;
						}
		                if(nikuData < minNiku)
		                {
							minNiku = nikuData;
						}
	                }

					if(sleeveNumber != "0" && sleeveNumber != "")
					{
						workSleeve++;
						//未登録なら追加
	                    if(list.Find(m => m.sleeveNumber == sleeveNumber).sleeveNumber != sleeveNumber)
		                {
	                        SleeveList sl = new SleeveList();
	                        sl.sleeveNumber= sleeveNumber;
							sl.shotCount = int.Parse(shotCount);
							list.Add(sl);
						}
		                else
		                {
							for(int j = 0; j < list.Count; j++)
							{
								if(list[j].sleeveNumber == sleeveNumber)
								{
									int sc = int.Parse(shotCount);
									if(list[j].shotCount < sc)
									{
										SleeveList tmpList = list[j];
										tmpList.shotCount = sc;
										list[j] = tmpList;
									}
								}
							}
						}
					}
	                
                }

            }

            //OK,NGの数を結合する
            okngStr = okCount.ToString() + "," + ngCount.ToString();
            //各不良原因の数を結合する
			ngcauseStr = nikuatsuNG + "," + 
						 kizuNG + "," + 
						 butsuNG + "," + 
						 yakeNG + "," + 
						 hibicrackNG + "," + 
						 gasukizuNG + "," + 
						 houshakizuNG + "," + 
						 giratsukikumoriNG + "," + 
						 hennikumendareNG + "," + 
						 yogoreNG + "," + 
						 hokoriNG + "," + 
						 keijoseidoNG + "," + 
						 etcNG;

			//標準偏差を求める
			double mean = sumNiku / nikuCount;
			double variance = (sumV_Niku / nikuCount) - (mean * mean);
			double stddev = Math.Sqrt(variance);

			//最大肉厚、最小肉厚、平均肉厚、標準偏差を結合する
			if(nikuCount > 0)
			{
				double aveNiku = (double)(sumNiku / nikuCount);
//				nikuInfo = maxNiku.ToString("F3") + "," + minNiku.ToString("F3") + "," + aveNiku.ToString("F3") + "," + nikuList.PopulationStandardDeviation().ToString("F3");
				nikuInfo = maxNiku.ToString("F3") + "," + minNiku.ToString("F3") + "," + aveNiku.ToString("F3") + "," + stddev.ToString("F3");
			}

			//スリーブ番号、ショット数を結合する
			int jj = 0;
			for(int i = (list.Count - 1); i >= 0; i--)
			{
				if(jj == 0)
				{
					sleeveInfo += list[i].sleeveNumber + "," + list[i].shotCount.ToString();
				}
				else
				{
					sleeveInfo += "," + list[i].sleeveNumber + "," + list[i].shotCount.ToString();
				}
				jj++;
			}

			//スリーブの稼働率
			kadou = (double)((double)workSleeve / (double)allSleeve);

        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = listView1.SelectedItems[0].Index;//上から0オリジンで数えた位置

            string detailData = "";
            int listlength = listView1.Items[index].SubItems.Count;
            for (int i = 0; i < listlength; i++)
            {
                if (i == 0)
                {
                    detailData = listView1.Items[index].SubItems[i].Text;
                    continue;
                }
                detailData += "," + listView1.Items[index].SubItems[i].Text;
            }

            Form5 form5 = new Form5();
            form5.SetDetailData(detailData);
            form5.ShowDialog();
        }

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("終了してよいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                e.Cancel = true;
            }
        }
    }
}
