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


			this.Text = "RedAnalyser Ver." + Version;
			this.label情報.Text = "CSVファイルをドラッグ＆ドロップしてください";
			this.groupBox1.Text = "読み込みCSVファイル";

			InitProgressGrid();
			InitMemberGrid();
			InitNoLimitTaskGrid();
		}

		//[Conditional("DEBUG")]
		//private void DebugControlView()
		//{
		//	this.textBox開発.Visible = true;
		//}

		private void GroupBoxProgressText()
		{
			this.groupBox3.Text = "期日ありタスク進捗情報 (" + DgvProgress.RowCount.ToString() + ")";
		}

		private void GroupBoxMemberText()
		{
			this.groupBox2.Text = "担当者別のタスク (" + DgvMember.RowCount.ToString() + ")";
		}

		private void GroupBoxNoLimitText()
		{
			this.groupBox4.Text = "期日未定のタスク (" + DgvNoLimitTask.RowCount.ToString() + ")";
		}

		private void InitProgressGrid()
		{
			GroupBoxProgressText();

			// ▼初期化中はコントロール使用不可
			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).BeginInit();

			this.DgvProgress.Init(this.groupBox3, "");

			this.DgvProgress.Columns.Clear();
			//this.DgvProgress.ScrollBars = ScrollBars.Vertical;

			// カラム(ヘッダ)の出力
			this.DgvProgress.Columns.Add("id", "#");
			this.DgvProgress.Columns["id"].Width = UseCsvTbl[CSV_TASK_ID];
			this.DgvProgress.Columns["id"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

			this.DgvProgress.Columns.Add("題名", "題名");
			this.DgvProgress.Columns["題名"].Width = UseCsvTbl[CSV_TASK_NAME];
			this.DgvProgress.Columns["題名"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

			this.DgvProgress.Columns.Add("担当者", "担当者");
			this.DgvProgress.Columns["担当者"].Width = UseCsvTbl[CSV_PERSON_NAME];
			this.DgvProgress.Columns["担当者"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

			// 「プログレスバー」項目を追加
			var pgb = new DataGridViewProgressBarColumn();
			//progressBar.DataPropertyName = "Progress";
			pgb.HeaderText = "進捗率";
			pgb.Name = "Progress";
			this.DgvProgress.Columns.Add(pgb);
			this.DgvProgress.Columns["Progress"].Width = UseCsvTbl[CSV_PROGRESS_BAR];
			this.DgvProgress.Columns["Progress"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

			this.DgvProgress.Columns.Add("期日", "期日");
			this.DgvProgress.Columns["期日"].Width = UseCsvTbl[CSV_DELIVERY_DAY];
			this.DgvProgress.Columns["期日"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;//※効果なし
			this.DgvProgress.Columns["期日"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;

			// 「残り日数」項目を追加
			this.DgvProgress.Columns.Add("残り日数", "残り");
			this.DgvProgress.Columns["残り日数"].Width = UseCsvTbl[CSV_REMAIMING];
			this.DgvProgress.Columns["残り日数"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;//※効果なし
			this.DgvProgress.Columns["残り日数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

			// ▲初期化が完了したら送信する
			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).EndInit();
		}

		private void InitMemberGrid()
		{
			GroupBoxMemberText();

			// ▼初期化中はコントロール使用不可
			((System.ComponentModel.ISupportInitialize)(this.DgvMember)).BeginInit();

			this.DgvMember.Init(this.groupBox2, "No.");

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
			pgb.HeaderText = "平均進捗率";
			this.DgvMember.Columns.Add(pgb);
			this.DgvMember.Columns["Progress"].Width = UseCsvTbl[CSV_PROGRESS_BAR];//75;
			this.DgvMember.Columns["Progress"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

			// ▲初期化が完了したら送信する
			((System.ComponentModel.ISupportInitialize)(this.DgvMember)).EndInit();
		}

		private void InitNoLimitTaskGrid()
		{
			GroupBoxNoLimitText();

			// ▼初期化中はコントロール使用不可
			((System.ComponentModel.ISupportInitialize)(this.DgvNoLimitTask)).BeginInit();

			this.DgvNoLimitTask.Init(this.groupBox4, "");

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


			// ▲初期化が完了したら送信する
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
			_TASK_NO = 0,
			_TITLE,
			_PERSON,

			_PROGRESS_BAR,//new

			_DELIVERY,
			_REMAINING,     // 残り日数※追加
			//_PROGRESS_BAR,
		}

		private void CsvReader()
		{
			var headerDic = new Dictionary<string, NameWidth>();
			var rowDicList = new List<Dictionary<string, string>>();
			var personsTask = new PersonsTask();

			using (StreamReader sr = new StreamReader(this.m_ReadFile, m_Encod)) {
				string line;
				for (int row = 0; (line = sr.ReadLine()) != null; row++) {
					string[] values = line.Split(',');
					// 横ライン分
					var dataDic = new Dictionary<string, string>();
					for (int column = 0; column < values.Length; column++) {
						if (this.UseCsvTbl[column] == 0) continue;

						// ヘッダの取得
						if (row == 0) {

						}
						// 本体データの取得
						else {
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
					}
					// 追加項目
					if (row != 0) {
						rowDicList.Add(dataDic);
					}
				}
			}

			this.DgvProgress.Rows.Clear();
			MakeProgressRow(rowDicList);

			MakePersonTaskGrid(personsTask);

			MakeNoLimitTaskGrid();
		}

		private void MakeProgressRow(List<Dictionary<string, string>> rowDicList)
		{
			NoLimitList = new List<Dictionary<string, string>>();

			int dicRowCount = 0;

			foreach (var dicCell in rowDicList) {
				string data;

				// 条件を満たしたタスクだけ表示
				if (dicCell.TryGetValue(CSV_PERSON_NAME.ToString(), out data)) {
					if (data == "\"\"") continue;
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

				//int cellCount = 0;
				foreach (var cell in dicCell) {
					bool b2 = dicCell.TryGetValue(cell.Key, out data);
					if (!b2) {
						MessageBox.Show("データが読込めない");
						continue;
					}
					switch (Convert.ToInt32(cell.Key)) {
						case CSV_TASK_ID:   // "#":   // MAKE_COLUM._TASK_NO:
							this.DgvProgress.Rows.Add(data);
							//cellCount++;
							break;

						case CSV_TASK_NAME:     // "題名":  // MAKE_COLUM._TITLE:
							this.DgvProgress.Rows[dicRowCount].Cells["題名"].Value = data;
							//cellCount++;
							break;

						case CSV_DELIVERY_DAY:  // "期日":  // MAKE_COLUM._DELIVERY:
							this.DgvProgress.Rows[dicRowCount].Cells["期日"].Value = data;
							//cellCount++;
							break;

						case CSV_PROGRESS_RATE: // "進捗率": // MAKE_COLUM._PROGRESS:
							//SetCell(data + "%");
							//cellCount++;
							break;

						case CSV_PERSON_NAME:   // "担当者": // MAKE_COLUM._PERSON:
							this.DgvProgress.Rows[dicRowCount].Cells["担当者"].Value = data;
							//cellCount++;
							break;

						case CSV_REMAIMING:  // "残り":  // MAKE_COLUM._REMAINING:
							DateTime dNow = DateTime.Now.Date;  // 時間なしの今日
							DateTime dTime = DateTime.Parse(dicCell[CSV_DELIVERY_DAY.ToString()]);
							TimeSpan span = dTime - dNow;
							if (span.Days <= 0) {
								this.DgvProgress.Rows[dicRowCount]
									.Cells[(int)MAKE_COLUM._REMAINING].Style.ForeColor = Color.Red;//赤文字
							}
							this.DgvProgress.Rows[dicRowCount].Cells["残り日数"].Value = span.Days.ToString() + "日";
							//cellCount++;
							break;

						case CSV_PROGRESS_BAR:
							// 「プログレスバー」の内容
							string progress;
							bool b3 = dicCell.TryGetValue(CSV_PROGRESS_RATE.ToString(), out progress);
							this.DgvProgress.Rows[dicRowCount].Cells[(int)MAKE_COLUM._PROGRESS_BAR].Value = Convert.ToInt32(progress);
							//cellCount++;
							break;
					}

					// ローカル関数
					//void SetCell(string value)
					//{
					//	this.DgvProgress.Rows[dicRowCount].Cells[cellCount].Value = value;
					//	cellCount++;
					//}
				}
				dicRowCount++;
			}
			GroupBoxProgressText();

		}

		private void MakeNoLimitTaskGrid()
		{
			this.DgvNoLimitTask.Rows.Clear();

			int row = 0;
			foreach (var dic in this.NoLimitList) {
				var aaa = dic[CSV_TASK_ID.ToString()];
				this.DgvNoLimitTask.Rows.Add(aaa);

				int cell = 1;
				string task = dic[CSV_TASK_NAME.ToString()];
				this.DgvNoLimitTask.Rows[row].Cells[cell++].Value = task;

				string name = dic[CSV_PERSON_NAME.ToString()];
				this.DgvNoLimitTask.Rows[row].Cells[cell++].Value = name;

				// 「プログレスバー」の内容
				string rate = dic[CSV_PROGRESS_RATE.ToString()];
				this.DgvNoLimitTask.Rows[row].Cells[cell++].Value = Convert.ToInt32(rate);

				row++;
			}
			GroupBoxNoLimitText();

		}

		private void MakePersonTaskGrid(PersonsTask personTask)
		{
			this.DgvMember.Rows.Clear();

			int row = 0;
			foreach (var pt in personTask.NameDic) {
				string name = pt.Key;
				this.DgvMember.Rows.Add(name);
				this.DgvMember.Rows[row].Cells[1].Value = pt.Value.Count.ToString();
				//this.DgvMember.Rows[row].Cells[2].Value = personTask.GetAverageProgress(name).ToString("F1") + "%";

				// 「プログレスバー」の内容
				float rate = personTask.GetAverageProgress(name);
				this.DgvMember.Rows[row].Cells[2].Value = Convert.ToInt32(rate);

				row++;
			}
			GroupBoxMemberText();

		}

		private class NameWidth
		{
			public int Width { get; set; }
			public string Name { get; set; }

			public NameWidth(string name, int width)
			{
				this.Name = name;
				this.Width = width;
			}
		}

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