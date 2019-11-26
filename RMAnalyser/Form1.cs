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

		public Form1()
		{
			InitializeComponent();

			this.Text = "RedAnalyser Ver." + Version;

			this.DgvProgress = new DGVProgress();
			// 左ヘッダをなくする(行番号を付けるイベントハンドラも削除すること)
			this.DgvProgress.RowHeadersVisible = false;

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
		};
		const int CSV_TASK_ID = 0;
		const int CSV_TASK_NAME = 6;
		const int CSV_PERSON_NAME = 8;
		const int CSV_DELIVERY_DAY = 13;
		const int CSV_PROGRESS_RATE = 15;
		const int CSV_REMAIMING = 20;

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

			// ▼初期化中はコントロール使用不可
			//((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).BeginInit();

			MakeProgressGrid(headerDic);
			MakeCellGrid(rowDicList);

			// ▲初期化が完了したら送信する
			//((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).EndInit();

		}

		private void MakeProgressGrid(Dictionary<string, NameWidth> headerDic)
		{
			// 左上隅のセルの値
			this.DgvProgress.TopLeftHeaderCell.Value = "タスク";
			this.DgvProgress.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

			// カラム(ヘッダ)の出力
			foreach (var h in headerDic) {
				string name = h.Value.Name;
				this.DgvProgress.Columns.Add(name, name);
				this.DgvProgress.Columns[name].Width = h.Value.Width;
			}
			// 項目を追加する
			this.DgvProgress.Columns.Add("残り日数", "残り");
			this.DgvProgress.Columns["残り日数"].Width = UseCsvTbl[CSV_REMAIMING];

			var progressBar = new DataGridViewProgressBarColumn();
			progressBar.DataPropertyName = "Progress";
			this.DgvProgress.Columns.Add(progressBar);



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

				int cellCount = 0;
				foreach (var cell in dicCell) {
					bool b2 = dicCell.TryGetValue(cell.Key, out data);
					int rowIndex = Convert.ToInt32(cell.Key);

					switch (rowIndex) {
						case CSV_TASK_ID:   // "#":   // MAKE_COLUM._TASK_NO:
							this.DgvProgress.Rows.Add(data);
							cellCount++;
							break;

						case CSV_TASK_NAME:     // "題名":  // MAKE_COLUM._TITLE:
						case CSV_DELIVERY_DAY:  // "期日":  // MAKE_COLUM._DELIVERY:
							this.DgvProgress.Rows[dicRowCount].Cells[cellCount].Value = data;
							cellCount++;
							break;

						case CSV_PROGRESS_RATE: // "進捗率": // MAKE_COLUM._PROGRESS:
							this.DgvProgress.Rows[dicRowCount].Cells[cellCount].Value = data + "%";
							cellCount++;
							break;

						case CSV_PERSON_NAME:   // "担当者": // MAKE_COLUM._PERSON:

							string name = "未割り当て";
							if (!data.Equals("\"\"")) {
								name = data;
							}
							this.DgvProgress.Rows[dicRowCount].Cells[cellCount].Value = name;
							cellCount++;
							break;

						case CSV_REMAIMING:  // "残り":  // MAKE_COLUM._REMAINING:
											 //dataGridView出力.Rows[dicRowCount].Cells[cell].Value = setRowData[cell] + "日";
											 //int span = Convert.ToInt32(data);
											 //if (span <= 0) {
											 //	dataGridView出力.Rows[dicRowCount]
											 //		.Cells[(int)MAKE_COLUM._REMAINING].Style.ForeColor = Color.Red;//赤文字
											 //}
							this.DgvProgress.Rows[dicRowCount].Cells[cellCount].Value = data + "日";
							cellCount++;
							break;
					}
				}
				dicRowCount++;
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