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
    public partial class Form1 : Form
    {

		public struct SleeveList
		{
			public string sleeveNumber;
			public int shotCount;
		}

		string selectedFile = "";

		string date = "";
		string time = "";
		string seikeikiName = "";
		string seihinName = "";

		string lastFileName = "";
		string sleeveNo = "";
		string z3Value = "";
		string ct1Value = "";
		string ct2Value = "";
		string cc1Value = "";
		string cc2Value = "";
		string cc3Value = "";
		string cc32Value = "";
		string cpValue = "";
		string tactValue = "";
		string Tkeisu = "";
		string z3hosei = "";
		string option1 = "";
		string option2 = "";
		string timeValue = "";
		string nikuUpLimit = "";
		string nikuData = "";
		string nikuLoLimit = "";
		string nikuResult = "";
		string resultCause = "";
		string currentHousharitsu = "";
		string currentOperator = "";
//		string sleeveNumber = "";
		string shotCount = "";
		string seikeiCount = "";
	
		string seikeikiHeader = "";
		string dstFile = "";

		FileStream fp;
		StreamReader sr = null;

		FormCalender formCalender = null;

        public Form1()
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
            ColumnHeader columnLslFileName = new ColumnHeader();
            ColumnHeader columnLslSleeveNo = new ColumnHeader();
            ColumnHeader columnZ3Value = new ColumnHeader();
            ColumnHeader columnCT1Value = new ColumnHeader();
            ColumnHeader columnCT2Value = new ColumnHeader();
            ColumnHeader columnCC1Value = new ColumnHeader();
            ColumnHeader columnCC2Value = new ColumnHeader();
            ColumnHeader columnCC3Value = new ColumnHeader();
            ColumnHeader columnCC32Value = new ColumnHeader();
            ColumnHeader columnCpValue = new ColumnHeader();
            ColumnHeader columnTact = new ColumnHeader();
            ColumnHeader columnTkeisu = new ColumnHeader();
            ColumnHeader columnZ3hosei = new ColumnHeader();
            ColumnHeader columnOption1 = new ColumnHeader();
            ColumnHeader columnOption2 = new ColumnHeader();
            ColumnHeader columnTimeStamp = new ColumnHeader();
            ColumnHeader columnNikuatsuUpper = new ColumnHeader();
            ColumnHeader columnNikuatsuValue = new ColumnHeader();
            ColumnHeader columnNikuatsuLower = new ColumnHeader();
			ColumnHeader columnResult = new ColumnHeader();
			ColumnHeader columnNgCause = new ColumnHeader();
			ColumnHeader columnHousharitsu = new ColumnHeader();
			ColumnHeader columnOperator = new ColumnHeader();
//			ColumnHeader columnResistSleeve = new ColumnHeader();
			ColumnHeader columnShotCount = new ColumnHeader();
			ColumnHeader columnSeikeiCount = new ColumnHeader();

			columnDate.Text = "日付";
			columnTime.Text = "時間";
			columnSeikeiki.Text = "成型機";
			columnSeihinName.Text = "製品名";
			columnLslFileName.Text = "ﾛｸﾞﾌｧｲﾙ名";
			columnLslSleeveNo.Text = "型No";
			columnZ3Value.Text = "Z3";
			columnCT1Value.Text = "初期温度";
			columnCT2Value.Text = "成型温度";
			columnCC1Value.Text = "CC1";
			columnCC2Value.Text = "CC2";
			columnCC3Value.Text = "CC3";
			columnCC32Value.Text = "CC3-2";
			columnCpValue.Text = "CP値";
			columnTact.Text = "ﾀｸﾄ";
			columnTkeisu.Text = "T係数";
			columnZ3hosei.Text = "Z3補正";
			columnOption1.Text = "OP1";
			columnOption2.Text = "OP2";
			columnTimeStamp.Text = "ﾀｲﾑｽﾀﾝﾌﾟ";
			columnNikuatsuUpper.Text = "肉厚上限";
			columnNikuatsuValue.Text = "肉厚測定値";
			columnNikuatsuLower.Text = "肉厚下限";
			columnResult.Text = "合否";
			columnNgCause.Text = "不良原因";
			columnHousharitsu.Text = "放射率";
			columnOperator.Text = "作業者";
//			columnResistSleeve.Text = "登録ｽﾘｰﾌﾞ";
			columnShotCount.Text = "ｼｮｯﾄ数";
			columnSeikeiCount.Text = "成型数";

            ColumnHeader[] colHeaderRegValue =
				{columnDate, columnTime, columnSeikeiki, columnSeihinName, columnLslFileName, columnLslSleeveNo, columnZ3Value, columnCT1Value, columnCT2Value, columnCC1Value,
				columnCC2Value, columnCC3Value, columnCC32Value, columnCpValue, columnTact, columnTkeisu, columnZ3hosei, columnOption1, columnOption2, columnTimeStamp,
				columnNikuatsuUpper, columnNikuatsuValue, columnNikuatsuLower, columnResult, columnNgCause, columnHousharitsu, columnOperator, /*columnResistSleeve, */columnShotCount, columnSeikeiCount};

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

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
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
			if(sr != null)
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

        private void Form1_Shown(object sender, EventArgs e)
        {
			//ListViewを一度クリアする
			listView1.Items.Clear();

			string[] item1 = {date, time, 
							  seikeikiName, seihinName, lastFileName, sleeveNo, z3Value, ct1Value, ct2Value, cc1Value, cc2Value, cc3Value,
							  cc32Value, cpValue, tactValue, Tkeisu, z3hosei, option1, option2, timeValue, nikuUpLimit, nikuData,
							  nikuLoLimit, nikuResult, resultCause, currentHousharitsu, currentOperator, /*sleeveNumber,*/ shotCount, seikeiCount};

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

				if(item1[23] == "OK")
				{
		            listView1.Items[0].BackColor = Color.Green;
				}
				else if(item1[23] == "NG")
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

					if(noonSta <= dt3 && dt3 < noonEnd)
					{
						listView1.Items[0].ForeColor = Color.Blue;//青
					}
					if(sunsetSta <= dt3 && dt3 < sunsetEnd)
					{
						listView1.Items[0].ForeColor = Color.Green;//緑
					}
					if(nightSta <= dt3 && dt3 < nightEnd)
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
	                this.Form1_Shown(sender, e);

	            }
			}
        }

        private void button2_Click(object sender, EventArgs e)
        {
			if(formCalender == null || formCalender.IsDisposed)
			{
				formCalender = new FormCalender();
                formCalender.Show();
				formCalender.SetParentForm(this, 1);
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
                    else if(j == 21)//肉厚測定値
					{
						nikuatsuSokutei = listitem;
					}
                    else if(j == 23)//合否
					{
						result = listitem;
					}
                    else if(j == 24)//不良原因
					{
						ngcause = listitem;
					}
                    else if(j == 5)//スリーブ番号
					{
						sleeveNumber = listitem;
					}
                    else if(j == 27)//ショット数
					{
						shotCount = listitem;
					}
                    else if(j == 28)//成型数
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
			for(int i = 0; i < listlength; i++)
			{
				if(i == 0)
				{
					detailData = listView1.Items[index].SubItems[i].Text;
					continue;
				}
				detailData += "," + listView1.Items[index].SubItems[i].Text;
			}

            Form3 form3 = new Form3();
			form3.SetDetailData(detailData);
            form3.ShowDialog();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("終了してよいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.No)
            {
                e.Cancel = true;
            }
        }
    }
}
