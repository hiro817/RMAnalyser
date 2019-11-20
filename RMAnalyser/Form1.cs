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



		public Form1()
		{
			InitializeComponent();

			this.Text = "RedAnalyser Ver." + Version;
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

			using (StreamReader sr = new StreamReader(this.m_ReadFile, m_Encod)) {

				string line;

				//while ((line = sr.ReadLine()) != null) {
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
							//m_HeaderNameWidth.Add(new HeaderNameWidth(this.NecessaryCsvTbl[column], values[column], column.ToString()));
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

				}
			}

			MakeHeaderRow(headerDic);

		}

		private void MakeHeaderRow(Dictionary<string, NameWidth> header)
		{
			//	// 
			//	// dataGridView出力
			//	// 
			//	this.dataGridView出力.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			//	this.dataGridView出力.Location = new System.Drawing.Point(8, 18);
			//	this.dataGridView出力.Name = "dataGridView出力";
			//	this.dataGridView出力.RowTemplate.Height = 21;
			//	this.dataGridView出力.Size = new System.Drawing.Size(583, 299);
			//	this.dataGridView出力.TabIndex = 0;

			var dgvProgress = new DGVProgress();
			// コントロールを使用できなくする
			((System.ComponentModel.ISupportInitialize)(dgvProgress)).BeginInit();

			// 左上隅のセルの値
			dgvProgress.TopLeftHeaderCell.Value = "タスク";
			//dgvProgress.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			dgvProgress.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

			// カラム(ヘッダ)の出力
			foreach (var h in header) {

				string name = h.Value.Name;
				int width = h.Value.Width;
				dgvProgress.Columns.Add(name, name);
				dgvProgress.Columns[name].Width = width;
			}

			this.groupBox3.Controls.Add(dgvProgress);
			dgvProgress.Location = new Point(10, 20);


			//dgvProgress.RowTemplate.Height = 21;//?

			// 自動でコントロールの四辺にドッキングして適切なサイズに調整される
			dgvProgress.Dock = DockStyle.Fill;
			//※↓手動でする場合
			//int width = this.groupBox3.Size.Width - 20;
			//int height = this.groupBox3.Size.Height - 30;
			//dgvProgress.Size = new Size(width, height);


			// 初期化が完了したら送信する
			((System.ComponentModel.ISupportInitialize)(dgvProgress)).EndInit();
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