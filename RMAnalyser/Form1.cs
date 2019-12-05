using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace RMAnalyser
{
	public partial class Form1 : Form
	{
		private readonly string Version = "1.00";
		private string m_ReadFile;
		private readonly Encoding m_Encod = Encoding.GetEncoding("Shift_JIS");

		private DGV DgvProgress = new DGV();
		private DGV DgvMember = new DGV();
		private DGV DgvNoLimitTask = new DGV();

		private List<Dictionary<string, string>> NoLimitList;

		private readonly string Nobady = "(未割り当て)";
		private readonly int[] UseCsvTbl = {
			45,		// 00 #(ID)		★CSV_TASK_ID
			0,		// 01 プロジェクト
			0,		// 02 トラッカー
			0,		// 03 親チケット
			0,		// 04 ステータス
			0,		// 05 優先度
			250,	// 06 題名		★CSV_TASK_NAME
			0,		// 07 作成者
			80,		// 08 担当者	★CSV_PERSON_NAME
			0,		// 09 更新日
			0,		// 10 カテゴリ
			0,		// 11 対象バージョン
			0,		// 12 開始日
			74,		// 13 期日		★CSV_DELIVERY_DAY
			0,		// 14 予定工数
			52,		// 15 進捗率	★CSV_PROGRESS_RATE
			0,		// 16 作成日
			0,		// 17 終了日
			0,		// 18 関連するチケット
			0,		// 19 プライベート
			// 以下追加分
			54,		// 20 残り日数	★CSV_REMAIMING
			75,		// 21 プログレスバー	★CSV_PROGRESS_BAR
		};

		//※下記はswitchで使う場合、readonlyには出来ない
		private const int CSV_TASK_ID = 0;
		private const int CSV_TASK_NAME = 6;
		private const int CSV_PERSON_NAME = 8;
		private const int CSV_DELIVERY_DAY = 13;
		private const int CSV_PROGRESS_RATE = 15;
		private const int CSV_REMAIMING = 20;
		private const int CSV_PROGRESS_BAR = 21;

		public Form1()
		{
			InitializeComponent();

			//this.CsvDic = Enum.GetNames(typeof(CSVDIC));
			//foreach (string e in this.CsvDic) {
			//	textBox開発.Text += e + "\r\n";
			//}

			this.Text = "RMAnalyser Ver." + Version;
			this.label情報.Text = "CSVファイルをドラッグ＆ドロップしてください";
			this.groupBox1.Text = "読み込みCSVファイル";

			InitDgvProgress();
			InitDgvMember();
			InitDgvNoLimitTask();

		}

		//[Conditional("DEBUG")]
		//private void DebugControlView()
		//{
		//	this.textBox開発.Visible = true;
		//}

		private void InitDgvProgress()
		{
			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).BeginInit();

			this.DgvProgress.SetGroupTextRowCount();

			this.DgvProgress.Init(this.groupBox3, "期日ありタスク進捗情報", "");

			this.DgvProgress.Columns.Clear();
			//this.DgvProgress.ScrollBars = ScrollBars.Vertical;//※常に垂直スクロールバーを表示させたい

			//this.DgvProgress.ColumnCount = 6;

			// カラム(ヘッダ)の出力
			var columns = new DataGridViewTextBoxColumn();//★
			columns.Name = "ID";
			columns.DataPropertyName = "Id";
			columns.HeaderText = "#";
			columns.Width = UseCsvTbl[CSV_TASK_ID];
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvProgress.Columns.Add(columns);

#if false
			columns = new DataGridViewTextBoxColumn();//★
			columns.Name = "TITLE";
			columns.DataPropertyName = "Title";
			columns.HeaderText = "題名";
			columns.Width = UseCsvTbl[CSV_TASK_NAME];
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvProgress.Columns.Add(columns);
#else
			this.DgvProgress.Columns.Add(
				new DataGridViewTextBoxColumn() {
					Name = "TITLE",
					DataPropertyName = "Title",
					HeaderText = "題名",
					Width = UseCsvTbl[CSV_TASK_NAME],
					//DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft,//何故か指定できない
				}
			);

#endif

			columns = new DataGridViewTextBoxColumn();//★
			columns.Name = "NAME";
			columns.DataPropertyName = "Name";
			columns.HeaderText = "名前";
			columns.Width = UseCsvTbl[CSV_PERSON_NAME];
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvProgress.Columns.Add(columns);

			// 「プログレスバー」項目を追加
			var pgb = new DataGridViewProgressBarColumn();
			pgb.Name = "BAR";
			pgb.DataPropertyName = "Progress";
			pgb.HeaderText = "進捗率";
			pgb.Width = UseCsvTbl[CSV_PROGRESS_BAR];
			pgb.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
			this.DgvProgress.Columns.Add(pgb);

			columns = new DataGridViewTextBoxColumn();//★
			columns.Name = "LIMIT";
			columns.DataPropertyName = "Limit";
			columns.HeaderText = "期日";
			columns.Width = UseCsvTbl[CSV_DELIVERY_DAY];
			columns.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;//※効果なし?
			this.DgvProgress.Columns.Add(columns);

			// 「残り日数」項目を追加
			columns = new DataGridViewTextBoxColumn();//★
			columns.Name = "REMAIMING";
			columns.DataPropertyName = "Remaiming";
			columns.HeaderText = "残り";
			columns.Width = UseCsvTbl[CSV_REMAIMING];
			columns.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;//※効果なし?
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
			this.DgvProgress.Columns.Add(columns);

			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).EndInit();
		}

		private void InitDgvMember()
		{
			((System.ComponentModel.ISupportInitialize)(this.DgvMember)).BeginInit();

			//this.DgvMember.Init(this.groupBox2, "担当者別のタスク", "No.");//※貼付け時に「No.」が不要になるので無くした＠19/12/03
			this.DgvMember.Init(this.groupBox2, "担当者別のタスク");

			// カラム(ヘッダ)の出力
			this.DgvMember.Columns.Add("担当者", "担当者");
			this.DgvMember.Columns["担当者"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvMember.Columns["担当者"].Width = UseCsvTbl[(int)CSV_PERSON_NAME];

			this.DgvMember.Columns.Add("タスク数", "数");
			this.DgvMember.Columns["タスク数"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
			this.DgvMember.Columns["タスク数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
			this.DgvMember.Columns["タスク数"].Width = 30;

			// 「プログレスバー」項目を追加
			var pgb = new DataGridViewProgressBarColumn();
			pgb.DataPropertyName = "Progress";
			pgb.Name = "Progress";
			//pgb.HeaderText = "平均進捗率";
			pgb.HeaderText = "平均";
			this.DgvMember.Columns.Add(pgb);
			this.DgvMember.Columns["Progress"].Width = UseCsvTbl[CSV_PROGRESS_BAR];//75;
			this.DgvMember.Columns["Progress"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

			((System.ComponentModel.ISupportInitialize)(this.DgvMember)).EndInit();
		}

		private void InitDgvNoLimitTask()
		{
			((System.ComponentModel.ISupportInitialize)(this.DgvNoLimitTask)).BeginInit();

			this.DgvNoLimitTask.Init(this.groupBox4, "期日未定のタスク", "");

			// カラム(ヘッダ)の出力
			this.DgvNoLimitTask.Columns.Add("id", "#");
			this.DgvNoLimitTask.Columns["id"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvNoLimitTask.Columns["id"].Width = UseCsvTbl[(int)CSV_TASK_ID];

			this.DgvNoLimitTask.Columns.Add("題名", "題名");
			this.DgvNoLimitTask.Columns["題名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvNoLimitTask.Columns["題名"].Width = UseCsvTbl[(int)CSV_TASK_NAME];

			this.DgvNoLimitTask.Columns.Add("担当者", "担当者");
			this.DgvNoLimitTask.Columns["担当者"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvNoLimitTask.Columns["担当者"].Width = UseCsvTbl[(int)CSV_PERSON_NAME];

			// 「プログレスバー」項目を追加
			var pgb = new DataGridViewProgressBarColumn();
			pgb.DataPropertyName = "Progress";
			pgb.HeaderText = "進捗率";
			pgb.Name = "Progress";
			this.DgvNoLimitTask.Columns.Add(pgb);
			this.DgvNoLimitTask.Columns["Progress"].Width = UseCsvTbl[CSV_PROGRESS_BAR];
			this.DgvNoLimitTask.Columns["Progress"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

			((System.ComponentModel.ISupportInitialize)(this.DgvNoLimitTask)).EndInit();
		}

		private void Form1_DragEnter(object sender, DragEventArgs e)
		{
			// 隠れていても前面にする
			this.Activate();
			// ドラッグされたデータがファイルならコピーする
			if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
				e.Effect = DragDropEffects.Copy;
			}
		}

		private void Form1_DragDrop(object sender, DragEventArgs e)
		{
			this.m_ReadFile = "";

			string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
			foreach (var file in files) {
				if (Path.GetExtension(file) == ".csv") {
					this.m_ReadFile = file;
					break;
				}
			}
			if (this.m_ReadFile != "") {
				this.label情報.Text = "";
				textBoxファイル名.Text = Path.GetFileName(this.m_ReadFile);// 拡張子ありのファイル名
				this.label情報.Text = "";

				CsvReader();
			}
		}

		private enum MAKE_COLUM
		{
			_ID = 0,
			_TITLE,
			_PERSON,
			_PROGRESS_BAR,//new
			_DELIVERY,
			_REMAINING,     // 残り日数※追加
			//_PROGRESS_BAR,
		}

		private void CsvReader()
		{
			//var headerDic = new Dictionary<string, NameWidth>();
			var rowDicList = new List<Dictionary<string, string>>();
			var personsTask = new PersonsTask();

			using (StreamReader sr = new StreamReader(this.m_ReadFile, m_Encod)) {
				string line;
				for (int row = 0; (line = sr.ReadLine()) != null; row++) {
					// カラム名をスキップ
					if (row == 0) continue;

					string[] values = line.Split(',');
					// 横ライン分
					var dataDic = new Dictionary<string, string>();
					for (int column = 0; column < values.Length; column++) {
						// 不要データの除外
						if (this.UseCsvTbl[column] == 0) continue;

						// 本体データの取得
						dataDic.Add(column.ToString(), values[column]);

						#region 担当者別の進捗情報の取得

						string progressName = this.Nobady;
						if (values[CSV_PERSON_NAME] != "\"\"") {
							progressName = values[CSV_PERSON_NAME];
						}
						int rate = Convert.ToInt32(values[CSV_PROGRESS_RATE]);
						if (!personsTask.IsNewPerson(progressName, rate)) {
							personsTask.AddProgress(progressName, rate);
						}

						#endregion 担当者別の進捗情報の取得
					}

					rowDicList.Add(dataDic);
				}
			}

			MakeProgressRow(rowDicList);
			MakePersonTaskGrid(personsTask);
			MakeNoLimitTaskGrid();
		}

		private void MakeProgressRow(List<Dictionary<string, string>> rowDicList)
		{
			this.DgvProgress.Rows.Clear();
			NoLimitList = new List<Dictionary<string, string>>();
			int dicRowCount = 0;

			foreach (var dicCell in rowDicList) {
				string data;

				// 条件を満たしたタスクだけ表示
				if (dicCell.TryGetValue(CSV_PERSON_NAME.ToString(), out data)) {
					if (data == "\"\"") {
						dicCell[CSV_PERSON_NAME.ToString()] = this.Nobady;
						this.NoLimitList.Add(dicCell);
						continue;
					}
				}
				if (dicCell.TryGetValue(CSV_PROGRESS_RATE.ToString(), out data)) {
					if (data == "100") continue;
				}
				if (dicCell.TryGetValue(CSV_DELIVERY_DAY.ToString(), out data)) {
					if (data == "\"\"") {
						this.NoLimitList.Add(dicCell);
						continue;
					}
				}

				// 残り日数の追加
				dicCell.Add(CSV_REMAIMING.ToString(), UseCsvTbl[CSV_REMAIMING].ToString());

				// プログレスバーの追加
				dicCell.Add(CSV_PROGRESS_BAR.ToString(), UseCsvTbl[CSV_PROGRESS_BAR].ToString());

#if true
				this.DgvProgress.Rows.Add();
				this.DgvProgress.Rows[dicRowCount].Cells["ID"].Value = dicCell[CSV_TASK_ID.ToString()];

				//this.DgvProgress.Rows.Add(dicCell[CSV_TASK_ID.ToString()]);


				this.DgvProgress.Rows[dicRowCount].Cells["TITLE"].Value = dicCell[CSV_TASK_NAME.ToString()];
				this.DgvProgress.Rows[dicRowCount].Cells["NAME"].Value = dicCell[CSV_PERSON_NAME.ToString()];
				this.DgvProgress.Rows[dicRowCount].Cells["BAR"].Value = Convert.ToInt32(dicCell[CSV_PROGRESS_RATE.ToString()]);
				this.DgvProgress.Rows[dicRowCount].Cells["LIMIT"].Value = dicCell[CSV_DELIVERY_DAY.ToString()];

				DateTime dNow = DateTime.Now.Date;  // 時間なしの今日
				DateTime dTime = DateTime.Parse(dicCell[CSV_DELIVERY_DAY.ToString()]);
				TimeSpan span = dTime - dNow;
				if (span.Days <= 0) {
					this.DgvProgress.Rows[dicRowCount].Cells[(int)MAKE_COLUM._REMAINING].Style.ForeColor = Color.Red;//赤文字
				}
				this.DgvProgress.Rows[dicRowCount].Cells["REMAIMING"].Value = span.Days.ToString() + "日";
#else

				foreach (var cell in dicCell) {
					bool b2 = dicCell.TryGetValue(cell.Key, out data);
					if (!b2) {
						MessageBox.Show("データが読込めない");
						continue;
					}
					switch (Convert.ToInt32(cell.Key)) {
						case CSV_TASK_ID:		// "#"
							this.DgvProgress.Rows.Add(data);
							break;

						case CSV_TASK_NAME:     // "題名"
							this.DgvProgress.Rows[dicRowCount].Cells["TITLE"].Value = data;
							break;

						//case CSV_PROGRESS_RATE: // "進捗率"はスキップ
						//	break;

						case CSV_PERSON_NAME:   // "担当者"
							this.DgvProgress.Rows[dicRowCount].Cells["NAME"].Value = data;
							break;

						case CSV_PROGRESS_BAR:  // プログレスバー
							string progress;
							bool b3 = dicCell.TryGetValue(CSV_PROGRESS_RATE.ToString(), out progress);
							this.DgvProgress.Rows[dicRowCount].Cells["BAR"].Value = Convert.ToInt32(progress);
							break;

						case CSV_DELIVERY_DAY:  // "期日"
							this.DgvProgress.Rows[dicRowCount].Cells["LIMIT"].Value = data;
							break;

						case CSV_REMAIMING:		 // "残り日"
							DateTime dNow = DateTime.Now.Date;  // 時間なしの今日
							DateTime dTime = DateTime.Parse(dicCell[CSV_DELIVERY_DAY.ToString()]);
							TimeSpan span = dTime - dNow;
							if (span.Days <= 0) {
								this.DgvProgress.Rows[dicRowCount]
									.Cells[(int)MAKE_COLUM._REMAINING].Style.ForeColor = Color.Red;//赤文字
							}
							this.DgvProgress.Rows[dicRowCount].Cells["REMAIMING"].Value = span.Days.ToString() + "日";
							break;
					}
				}
#endif
				dicRowCount++;
			}
			this.DgvProgress.SetGroupTextRowCount();
		}

		private void MakePersonTaskGrid(PersonsTask personTask)
		{
			this.DgvMember.Rows.Clear();

			int row = 0;
			foreach (var pt in personTask.NameDic) {
				string name = pt.Key;
				this.DgvMember.Rows.Add(name);
				this.DgvMember.Rows[row].Cells[1].Value = pt.Value.Count.ToString();

				// 「プログレスバー」の内容
				float rate = personTask.GetAverageProgress(name);
				this.DgvMember.Rows[row].Cells[2].Value = Convert.ToInt32(rate);
				//this.DgvMember.Rows[row].Cells[2].Value = personTask.GetAverageProgress(name).ToString("F1") + "%";

				row++;
			}
			this.DgvMember.SetGroupTextRowCount();

		}

		private void MakeNoLimitTaskGrid()
		{
			this.DgvNoLimitTask.Rows.Clear();

			int row = 0;
			foreach (var dic in this.NoLimitList) {
				this.DgvNoLimitTask.Rows.Add(dic[CSV_TASK_ID.ToString()]);
				this.DgvNoLimitTask.Rows[row].Cells["題名"].Value = dic[CSV_TASK_NAME.ToString()];
				this.DgvNoLimitTask.Rows[row].Cells["担当者"].Value = dic[CSV_PERSON_NAME.ToString()];
				this.DgvNoLimitTask.Rows[row].Cells["Progress"].Value = Convert.ToInt32(dic[CSV_PROGRESS_RATE.ToString()]);
				row++;
			}
			this.DgvNoLimitTask.SetGroupTextRowCount();
		}

		//private class NameWidth
		//{
		//	public int Width { get; set; }
		//	public string Name { get; set; }

		//	public NameWidth(string name, int width)
		//	{
		//		this.Name = name;
		//		this.Width = width;
		//	}
		//}

		private void button担当者別_Click(object sender, EventArgs e)
		{
			this.label情報.Text = "担当者別のタスクを"+this.DgvMember.ButtonClick();
		}

		private void button期日あり進捗_Click(object sender, EventArgs e)
		{
			this.label情報.Text = "期日あり進捗のタスクを" + this.DgvProgress.ButtonClick();
		}
		private void button期日未定タスク_Click(object sender, EventArgs e)
		{
			this.label情報.Text = "期日未定タスクを" + this.DgvNoLimitTask.ButtonClick();
		}

	}

}

