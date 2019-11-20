using System.Collections.Generic;
using System.IO;
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
			var headerDic = new Dictionary<string, string>();
			var cellRowList = new List<string>();

			using (StreamReader sr = new StreamReader(this.m_ReadFile, m_Encod)) {

				string line;

				//while ((line = sr.ReadLine()) != null) {
				for (int row = 0; (line = sr.ReadLine()) != null; row++) {
					string[] values = line.Split(',');

					// 横ライン分
					var dic = new Dictionary<string, string>();

					for (int column = 0; column < values.Length; column++) {
						// 必要なデータだけ取り込む
						if (this.UseCsvTbl[column] == 0) continue;

						// ヘッダ
						if (row == 0) {
							// データから取り込んでヘッダを作成
							//m_HeaderNameWidth.Add(new HeaderNameWidth(this.NecessaryCsvTbl[column], values[column], column.ToString()));
							headerDic.Add(column.ToString(), values[column]);
						}
						// 本体
						else {
							string value = values[column];
							dic.Add(column.ToString(), value);
							cellRowList.Add(value);
						}
					}
					// 追加項目

				}
			}

			//※確認用
			foreach (var d in headerDic) {
				textBox開発.Text += d.Value + "\t";
			}

		}
	}
}
