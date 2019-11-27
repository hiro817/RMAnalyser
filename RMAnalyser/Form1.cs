using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RMAnalyser
{
	public partial class Form1 : Form
	{
		private readonly string Version = "1.00";
		private string m_ReadFile;
		readonly Encoding m_Encod = Encoding.GetEncoding("Shift_JIS");

		private DGVProgress DgvProgress;
		private DGVProgress DgvMember;


		public Form1()
		{
			InitializeComponent();

			this.Text = "RedAnalyser Ver." + Version;

			this.label情報.Text = "CSVファイルをドラッグ＆ドロップしてください";

			this.groupBox1.Text = "読み込みCSVファイル";
			this.groupBox2.Text = "担当者別タスク";
			this.groupBox3.Text = "納期別タスク";

			ProgressGridInit();
			MemberGridInit();

		}

		private void ProgressGridInit()
		{
			this.DgvProgress = new DGVProgress();
			this.DgvProgress.Init();

			// ▼初期化中はコントロール使用不可
			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).BeginInit();
			this.DgvProgress.Init();
			// 左ヘッダをなくする(行番号を付けるイベントハンドラも削除すること)
			this.DgvProgress.RowHeadersVisible = false;
			// ▲初期化が完了したら送信する
			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).EndInit();
		}

		private void MemberGridInit()
		{

			this.DgvMember = new DGVProgress();
			this.DgvMember.Init();

			// 左上隅のセルの値
			this.DgvMember.TopLeftHeaderCell.Value = "担当";
			this.DgvMember.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

			// カラム(ヘッダ)の出力
			this.DgvMember.Columns.Add("担当者", "担当者");
			this.DgvMember.Columns["担当者"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvMember.Columns["担当者"].Width = UseCsvTbl[(int)CSV_PERSON_NAME];

			this.DgvMember.Columns.Add("タスク数", "数");
			this.DgvMember.Columns["タスク数"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
			this.DgvMember.Columns["タスク数"].Width = 30;// NecessaryCsvTbl[(int)CSV_PROGRESS_RATE];
			this.DgvMember.Columns["タスク数"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

			this.DgvMember.Columns.Add("平均進捗率", "平均");
			this.DgvMember.Columns["平均進捗率"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
			this.DgvMember.Columns["平均進捗率"].Width = UseCsvTbl[(int)CSV_PROGRESS_RATE];
			this.DgvMember.Columns["平均進捗率"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

			// GroupBoxに追加する位置を設定
			this.groupBox2.Controls.Add(this.DgvMember);
			this.DgvMember.Location = new Point(10, 20);
			// 自動でコントロールの四辺にドッキングして適切なサイズに調整される
			this.DgvMember.Dock = DockStyle.Fill;
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

		readonly int[] UseCsvTbl = {
			45,		// 00 #(ID)		★CSV_TASK_ID
			0,		// 01 プロジェクト
			0,		// 02 トラッカー
			0,		// 03 親チケット
			0,		// 04 ステータス
			0,		// 05 優先度
			250,	// 06 題名		★CSV_TASK_NAME
			0,		// 07 作成者
			85,		// 08 担当者	★CSV_PERSON_NAME
			0,		// 09 更新日
			0,		// 10 カテゴリ
			0,		// 11 対象バージョン
			0,		// 12 開始日
			74,		// 13 期日		★CSV_DELIVERY_DAY
			0,		// 14 予定工数
			58,		// 15 進捗率	★CSV_PROGRESS_RATE
			0,		// 16 作成日
			0,		// 17 終了日
			0,		// 18 関連するチケット
			0,		// 19 プライベート
			// 以下追加分
			54,		// 20 残り日数	★CSV_REMAIMING
			75,		// 21
		};
		const int CSV_TASK_ID = 0;
		const int CSV_TASK_NAME = 6;
		const int CSV_PERSON_NAME = 8;
		const int CSV_DELIVERY_DAY = 13;
		const int CSV_PROGRESS_RATE = 15;
		const int CSV_REMAIMING = 20;
		const int CSV_PROGRESS_BAR = 21;

		enum MAKE_COLUM
		{
			_TASK_NO = 0,
			_TITLE,
			_PERSON,
			_DELIVERY,
			_PROGRESS,
			_REMAINING,     // 残り日数※追加
		}

		private void CsvReader()
		{
			var headerDic = new Dictionary<string, NameWidth>();
			var rowDicList = new List<Dictionary<string, string>>();

			PersonsTask personsTask = new PersonsTask();

			using (StreamReader sr = new StreamReader(this.m_ReadFile, m_Encod)) {
				string line;
				for (int row = 0; (line = sr.ReadLine()) != null; row++) {
					string[] values = line.Split(',');
					// 横ライン分
					var dataDic = new Dictionary<string, string>();
					for (int column = 0; column < values.Length; column++) {
						// 必要なデータだけ取り込む
						if (this.UseCsvTbl[column] == 0) continue;

						// ヘッダ
						if (row == 0) {
							// データから取り込んでヘッダを作成
							var nameWidth = new NameWidth(values[column], this.UseCsvTbl[column]);
							headerDic.Add(column.ToString(), nameWidth);
						}
						// 本体
						else {
							string value = values[column];
							dataDic.Add(column.ToString(), value);

							#region 担当者別の進捗情報の取得
							string progressName = (values[CSV_PERSON_NAME] != "\"\"") ?
								values[CSV_PERSON_NAME] : "(未割り当て)";

							int rate = Convert.ToInt32(values[CSV_PROGRESS_RATE]);
							if (!personsTask.IsNewPerson(progressName, rate))
							{
								personsTask.AddProgress(progressName, rate);
							}
							#endregion

						}
					}
					// 追加項目
					if (row != 0) {
						rowDicList.Add(dataDic);
					}
				}
			}

			this.DgvProgress.Rows.Clear();
			this.DgvProgress.Columns.Clear();

			this.DgvMember.Rows.Clear();

			MakeProgressGrid(headerDic);
			MakeCellGrid(rowDicList);

			MakePersonTaskGrid(personsTask);

		}

		private void MakeProgressGrid(Dictionary<string, NameWidth> headerDic)
		{
			// 左上隅のセルの値
			this.DgvProgress.TopLeftHeaderCell.Value = "タスク";
			this.DgvProgress.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

			// カラム(ヘッダ)の出力
			foreach (var h in headerDic)
			{
				string name = h.Value.Name;
				this.DgvProgress.Columns.Add(name, name);
				this.DgvProgress.Columns[name].Width = h.Value.Width;
			}
			// 項目を追加
			this.DgvProgress.Columns.Add("残り日数", "残り");
			this.DgvProgress.Columns["残り日数"].Width = UseCsvTbl[CSV_REMAIMING];

			// 項目を追加
			var progressBar = new DataGridViewProgressBarColumn();
			progressBar.DataPropertyName = "Progress";
			progressBar.HeaderText = "Progress";
			progressBar.Name = "Progress";
			this.DgvProgress.Columns.Add(progressBar);

			this.DgvProgress.Columns["Progress"].Width = UseCsvTbl[CSV_PROGRESS_BAR];

			// GroupBoxに追加する位置を設定
			this.groupBox3.Controls.Add(this.DgvProgress);
			this.DgvProgress.Location = new Point(10, 20);
			// 自動でコントロールの四辺にドッキングして適切なサイズに調整される
			this.DgvProgress.Dock = DockStyle.Fill;
		}

		private void MakeCellGrid(List<Dictionary<string, string>> list)
		{
			int dicRowCount = 0;
			for (int column = 0; column < list.Count; column++) {
				var dicCell = list[column];
				string data;

				// 条件を満たしたタスクだけ表示
				if (dicCell.TryGetValue(CSV_PERSON_NAME.ToString(), out data)) {
					if (data == "\"\"") continue;
				}
				if (dicCell.TryGetValue(CSV_PROGRESS_RATE.ToString(), out data)) {
					if (data == "100") continue;
				}
				if (dicCell.TryGetValue(CSV_DELIVERY_DAY.ToString(), out data)) {
					if (data == "\"\"") continue;
				}

				// 残り日数の算出と追加
				dicCell.Add(CSV_REMAIMING.ToString(), "99");


				dicCell.Add(CSV_PROGRESS_BAR.ToString(), "60");


				int cellCount = 0;
				foreach (var cell in dicCell) {
					bool b2 = dicCell.TryGetValue(cell.Key, out data);
					if (!b2) {
						MessageBox.Show("データが読込めない");
						continue;
					}
					switch (Convert.ToInt32(cell.Key)) {
						case CSV_TASK_ID:   // "#":   // MAKE_COLUM._TASK_NO:
							this.DgvProgress.Rows.Add(data);
							cellCount++;
							break;

						case CSV_TASK_NAME:     // "題名":  // MAKE_COLUM._TITLE:
						case CSV_DELIVERY_DAY:  // "期日":  // MAKE_COLUM._DELIVERY:
							SetCell(data);
							break;

						case CSV_PROGRESS_RATE: // "進捗率": // MAKE_COLUM._PROGRESS:
							SetCell(data + "%");
							break;

						case CSV_PERSON_NAME:   // "担当者": // MAKE_COLUM._PERSON:

							string name = "未割り当て";
							if (!data.Equals("\"\"")) {
								name = data;
							}
							SetCell(name);
							break;

						case CSV_REMAIMING:  // "残り":  // MAKE_COLUM._REMAINING:
											 // 時間なしの今日
							DateTime dNow = DateTime.Now.Date;

							DateTime dTime = DateTime.Parse(dicCell[CSV_DELIVERY_DAY.ToString()]);
							TimeSpan span = dTime - dNow;

							if (span.Days <= 0) {
								this.DgvProgress.Rows[dicRowCount]
									.Cells[(int)MAKE_COLUM._REMAINING].Style.ForeColor = Color.Red;//赤文字
							}
							SetCell(span.Days.ToString() + "日");
							break;

						case CSV_PROGRESS_BAR:
							SetCell("10");
							break;
					}


					// ローカル関数
					void SetCell(string value)
					{
						this.DgvProgress.Rows[dicRowCount].Cells[cellCount].Value = value;
						cellCount++;
					}

				}
				dicRowCount++;
			}
		}

		private void MakePersonTaskGrid(PersonsTask personTask)
		{

			int row = 0;
			foreach (var pt in personTask.NameDic)
			{
				string name = pt.Key;
				this.DgvMember.Rows.Add(name);  // 0

				this.DgvMember.Rows[row].Cells[1].Value = pt.Value.Count.ToString();

				this.DgvMember.Rows[row].Cells[2].Value = personTask.GetAverageProgress(name).ToString("F1") + "%";

				row++;

			}

		}


		class NameWidth
		{
			public int Width { get; set; }
			public string Name { get; set; }

			public NameWidth(string name, int width)
			{
				this.Name = name;
				this.Width = width;
			}
		}

	}
}