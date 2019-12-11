#define SW_PARENT

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
		private readonly string Version = "2.20";
		/*
			Ver.2.20	19/12/11	親チケットも表示・EXCELで開いている場合のエラーを追加
			Ver.2.10	19/12/09	担当者の（未割り当て）が先頭になるようにソート
						19/12/11	(未割り当て)を(親チケット)に変更
			Ver.2.00	19/12/06	DGVの内容の進捗率には％を付けてクリップボードにコピー
						19/12/09	「すべて項目」のCSVだけ処理するように変更
			Ver.1.20	19/12/05	期日ありのDGVの項目を変更（＃と題名を右側に移動）／進捗率に％を追加
			Ver.1.10	19/12/05	担当者別タスク数が間違っていた／期日未定タスクに（未割り当て）も表示
		*/

		private string m_ReadFile;
		private readonly Encoding m_Encod = Encoding.GetEncoding("Shift_JIS");

		private DGV DgvProgress = new DGV();
		private DGV DgvMember = new DGV();
		private DGV DgvNoLimitTask = new DGV();
		private List<Dictionary<int, string>> NoLimitList;
		private List<string> ParentList;

		private readonly string ReadableCsvWord = "#,プロジェクト,トラッカー,親チケット,ステータス,優先度,題名,作成者,担当者,更新日,カテゴリ,対象バージョン,開始日,期日,予定工数,進捗率,作成日,終了日,関連するチケット,プライベート";
		private readonly string Nobady = "(親チケット)";

		private readonly int[] UseCsvTbl = {
			45,		// 00 #(ID)		★CSV_TASK_ID
			0,		// 01 プロジェクト
			0,      // 02 トラッカー
			45,		// 03 親チケット★CSV_PARENT
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

		private const int CSV_PARENT = 3;
		private const int CSV_TASK_NAME = 6;
		private const int CSV_PERSON_NAME = 8;
		private const int CSV_DELIVERY_DAY = 13;
		private const int CSV_PROGRESS_RATE = 15;
		private const int CSV_REMAIMING = 20;
		private const int CSV_PROGRESS_BAR = 21;

		public Form1()
		{
			InitializeComponent();

			this.Text = "RMAnalyser Ver." + Version;
			this.label情報.Text = "CSVファイルをドラッグ＆ドロップしてください";
			this.groupBox1.Text = "読み込みCSVファイル";
			this.textBox開発.Text = "進捗率100%は含まれていません\r\n";

			InitDgvProgress();
			InitDgvMember();
			InitDgvNoLimitTask();
		}

		//[Conditional("DEBUG")]
		//private void DebugControlView()
		//{
		//	this.textBox開発.Visible = true;
		//}

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

		private void CsvReader()
		{
			var rowDicList = new List<Dictionary<int, string>>();
			var personsTask = new PersonsTask();
			this.ParentList = new List<string>();

			try {
				using (StreamReader sr = new StreamReader(this.m_ReadFile, m_Encod)) {
					string line;
					for (int row = 0; (line = sr.ReadLine()) != null; row++) {
						// カラム名をスキップ
						if (row == 0) {
							if (line != this.ReadableCsvWord) {
								MessageBox.Show(
									"RedmineのCSVは「すべての項目」を選択したファイルを使用してください。",
									"エラー",
									MessageBoxButtons.OK,
									MessageBoxIcon.Error);
								return;
							}
							continue;
						}
						string[] values = line.Split(',');
						// 横ライン分
						var dataDic = new Dictionary<int, string>();
						for (int column = 0; column < values.Length; column++) {
							// 不要データの除外
							if (this.UseCsvTbl[column] == 0) continue;

							// 本体データの取得
							dataDic.Add(column, values[column]);

							#region 担当者別の進捗情報の取得

							if (column == CSV_PERSON_NAME) {
								string progressName = this.Nobady;
								if (values[CSV_PERSON_NAME] != "\"\"") {
									progressName = values[CSV_PERSON_NAME];
								}
								else {
									progressName = this.Nobady;
								}

								int rate = Convert.ToInt32(values[CSV_PROGRESS_RATE]);
								if (!personsTask.IsNewPerson(progressName, rate)) {
									personsTask.AddProgress(progressName, rate);
								}
							}

							#endregion 担当者別の進捗情報の取得

							if (column == CSV_PARENT) {
								this.ParentList.Add(values[CSV_PARENT]);
							}
						}

						rowDicList.Add(dataDic);
					}
				}
			}
			catch (Exception e) {
				MessageBox.Show("Excelで開いている場合は閉じてから使ってください",
					"エラー", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
			}
			MakeProgressRow(rowDicList);
			MakePersonTaskGrid(personsTask);
			MakeNoLimitTaskGrid();
		}

		private void InitDgvProgress()
		{
			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).BeginInit();

			this.DgvProgress.SetGroupTextRowCount();
			this.DgvProgress.Init(this.groupBox3, "期日ありタスク進捗情報", "");

			this.DgvProgress.Columns.Clear();
			//this.DgvProgress.ScrollBars = ScrollBars.Vertical;//※横幅を一定にするため常に垂直スクロールバーを表示させたい

			// カラム(ヘッダ)の出力
			DataGridViewTextBoxColumn columns;

			columns = new DataGridViewTextBoxColumn();//★
			columns.HeaderText = "名前";
			columns.Name = "NAME";
			columns.DataPropertyName = "Name";
			columns.Width = UseCsvTbl[CSV_PERSON_NAME];
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvProgress.Columns.Add(columns);

			// 「プログレスバー」項目を追加
			var pgb = new DataGridViewProgressBarColumn();
			pgb.HeaderText = "進捗率";
			pgb.Name = "BAR";
			pgb.DataPropertyName = "Progress";
			pgb.Width = UseCsvTbl[CSV_PROGRESS_BAR];
			pgb.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
			this.DgvProgress.Columns.Add(pgb);

			columns = new DataGridViewTextBoxColumn();//★
			columns.HeaderText = "期日";
			columns.Name = "LIMIT";
			columns.DataPropertyName = "Limit";
			columns.Width = UseCsvTbl[CSV_DELIVERY_DAY];
			columns.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;//※効果なし?
			this.DgvProgress.Columns.Add(columns);

			// 「残り日数」項目を追加
			columns = new DataGridViewTextBoxColumn();//★
			columns.HeaderText = "残り";
			columns.Name = "REMAIMING";
			columns.DataPropertyName = "Remaiming";
			columns.Width = UseCsvTbl[CSV_REMAIMING];
			columns.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;//※効果なし?
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
			this.DgvProgress.Columns.Add(columns);

			columns = new DataGridViewTextBoxColumn();//★
			columns.HeaderText = "#";
			columns.Name = "ID";
			columns.DataPropertyName = "Id";
			columns.Width = UseCsvTbl[CSV_TASK_ID];
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvProgress.Columns.Add(columns);

			columns = new DataGridViewTextBoxColumn();//★
			columns.HeaderText = "親#";
			columns.Name = "PARENT";
			columns.DataPropertyName = "Parent";
			columns.Width = UseCsvTbl[CSV_PARENT];
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvProgress.Columns.Add(columns);

			columns = new DataGridViewTextBoxColumn();//★
			columns.HeaderText = "題名";
			columns.Name = "TITLE";
			columns.DataPropertyName = "Title";
			columns.Width = UseCsvTbl[CSV_TASK_NAME];
			columns.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
			this.DgvProgress.Columns.Add(columns);

			((System.ComponentModel.ISupportInitialize)(this.DgvProgress)).EndInit();
		}

		private void MakeProgressRow(List<Dictionary<int, string>> rowDicList)
		{
			this.DgvProgress.Rows.Clear();
			NoLimitList = new List<Dictionary<int, string>>();
			int dicRowCount = 0;

			//名前順にソート→ （親チケット）が先頭を先頭にする
			var sort = rowDicList
				.OrderBy(n => n[CSV_PERSON_NAME].ToString())
				.ThenBy(n => n[CSV_TASK_ID].ToString())
				.ToList();
			foreach (var dicCell in sort) {
				// 条件を満たしたタスクだけ表示
				if (dicCell[CSV_PROGRESS_RATE] == "100") continue;

				string name = dicCell[CSV_PERSON_NAME];
				string day = dicCell[CSV_DELIVERY_DAY];

				if (name == "\"\"") {
					dicCell[CSV_PERSON_NAME] = this.Nobady;
				}

				if (name == "\"\"" && day == "\"\"") {
					dicCell[CSV_PERSON_NAME] = this.Nobady;
					this.NoLimitList.Add(dicCell);
					continue;
				}
				else
				if (day == "\"\"") {
					this.NoLimitList.Add(dicCell);
					continue;
				}

				// 残り日数の追加
				dicCell.Add(CSV_REMAIMING, UseCsvTbl[CSV_REMAIMING].ToString());

				// プログレスバーの追加
				dicCell.Add(CSV_PROGRESS_BAR, UseCsvTbl[CSV_PROGRESS_BAR].ToString());

				this.DgvProgress.Rows.Add();
				string myId = dicCell[CSV_TASK_ID];
				this.DgvProgress.Rows[dicRowCount].Cells["ID"].Value = myId;
				if (this.ParentList.Any(id => id == myId)) {
					this.DgvProgress.Rows[dicRowCount].Cells["ID"].Style.BackColor = Color.Yellow;
				}

				string parent = dicCell[CSV_PARENT];
				if (parent == "\"\"") {
					parent = String.Empty;
				}
				else {
					this.DgvProgress.Rows[dicRowCount].Cells["PARENT"].Style.ForeColor = Color.Red;
				}
				this.DgvProgress.Rows[dicRowCount].Cells["PARENT"].Value = parent;

				this.DgvProgress.Rows[dicRowCount].Cells["TITLE"].Value = dicCell[CSV_TASK_NAME];
				this.DgvProgress.Rows[dicRowCount].Cells["NAME"].Value = dicCell[CSV_PERSON_NAME];
				this.DgvProgress.Rows[dicRowCount].Cells["BAR"].Value = Convert.ToInt32(dicCell[CSV_PROGRESS_RATE]);
				this.DgvProgress.Rows[dicRowCount].Cells["LIMIT"].Value = dicCell[CSV_DELIVERY_DAY];

				DateTime dNow = DateTime.Now.Date;  // 時間なしの今日
				DateTime dTime = DateTime.Parse(dicCell[CSV_DELIVERY_DAY]);
				TimeSpan span = dTime - dNow;
				this.DgvProgress.Rows[dicRowCount].Cells["REMAIMING"].Value = span.Days.ToString() + "日";
				if (span.Days <= 0) {
					this.DgvProgress.Rows[dicRowCount].Cells["REMAIMING"].Style.ForeColor = Color.Red;//赤文字
				}
				dicRowCount++;
			}
			this.DgvProgress.SetGroupTextRowCount();
		}

		private void InitDgvMember()
		{
			((System.ComponentModel.ISupportInitialize)(this.DgvMember)).BeginInit();

			this.DgvMember.Columns.Clear();

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
			pgb.HeaderText = "平均進捗率";
			//pgb.HeaderText = "平均";↑でも大丈夫だった＠19/12/09
			this.DgvMember.Columns.Add(pgb);
			this.DgvMember.Columns["Progress"].Width = UseCsvTbl[CSV_PROGRESS_BAR];//75;
			this.DgvMember.Columns["Progress"].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

			((System.ComponentModel.ISupportInitialize)(this.DgvMember)).EndInit();
		}

		private void MakePersonTaskGrid(PersonsTask personTask)
		{
			this.DgvMember.Rows.Clear();
			int row = 0;
			//名前順にソート→ （親チケット）を先頭にする
			var sort = new SortedDictionary<string, List<float>>(personTask.NameDic);
			foreach (var pt in sort) {
				string name = pt.Key;
				// 担当者
				this.DgvMember.Rows.Add(name);
				// タスク数
				this.DgvMember.Rows[row].Cells[1].Value = pt.Value.Count.ToString();
				// 進捗率
				this.DgvMember.Rows[row].Cells[2].Value = Convert.ToInt32(personTask.GetAverageProgress(name));
				row++;
			}
			this.DgvMember.SetGroupTextRowCount();
		}

		private void InitDgvNoLimitTask()
		{
			((System.ComponentModel.ISupportInitialize)(this.DgvNoLimitTask)).BeginInit();

			this.DgvNoLimitTask.Columns.Clear();

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

		private void MakeNoLimitTaskGrid()
		{
			this.DgvNoLimitTask.Rows.Clear();

			int row = 0;

			//名前順にソート→ （親チケット）を先頭にする
			var sort = this.NoLimitList
				.OrderBy(n => n[CSV_PERSON_NAME].ToString())
				.ThenBy(n => n[CSV_TASK_ID].ToString())
				.ToList();
			foreach (var dic in sort) {
				this.DgvNoLimitTask.Rows.Add(dic[CSV_TASK_ID]);
				this.DgvNoLimitTask.Rows[row].Cells["題名"].Value = dic[CSV_TASK_NAME];
				this.DgvNoLimitTask.Rows[row].Cells["担当者"].Value = dic[CSV_PERSON_NAME];
				this.DgvNoLimitTask.Rows[row].Cells["Progress"].Value = Convert.ToInt32(dic[CSV_PROGRESS_RATE]);
				row++;
			}
			this.DgvNoLimitTask.SetGroupTextRowCount();
		}

		private void button担当者別_Click(object sender, EventArgs e)
		{
			this.label情報.Text = "担当者別タスクをクリップボードにコピーしました！";
			this.DgvMember.CopyToClipboard(2);
		}

		private void button期日あり進捗_Click(object sender, EventArgs e)
		{
			this.label情報.Text = "期日あり進捗タスクをクリップボードにコピーしました！";
			this.DgvProgress.CopyToClipboard(1);
		}

		private void button期日未定タスク_Click(object sender, EventArgs e)
		{
			this.label情報.Text = "期日未定タスクをクリップボードにコピーしました！";
			this.DgvNoLimitTask.CopyToClipboard(3);
		}
	}
}