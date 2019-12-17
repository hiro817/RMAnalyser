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
		private readonly string Version = "2.30";
		/*
						19/12/17	GitHubにPrivateで追加
			Ver.2.30	19/12/13	特定のカラムを右寄せに出来るようにした
						19/12/12	内部処理（クラスの分割）/クラス追加
									例外処理の追加
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
		private readonly string Nobody = "(親チケット)";

		private readonly ColumnHeader[] ProgressHeaderTbl = {
			new ColumnHeader("NAME",        "名前",    UseCsv._PERSON_NAME),
			new ColumnHeader( "BAR",        "進捗率",  UseCsv._PROGRESS_BAR, true),
			new ColumnHeader("LIMIT",       "期日",    UseCsv._DELIVERY_DAY),
			new ColumnHeader( "REMAIMING",  "残り",    UseCsv._REMAIMING_R),	//※右寄せ
			new ColumnHeader( "ID",         "#",       UseCsv._TASK_ID),
			new ColumnHeader( "PARENT",     "親#",     UseCsv._PARENT),
			new ColumnHeader( "TITLE",      "題名",    UseCsv._TASK_NAME),
		};

		private readonly ColumnHeader[] MemberHeaderTbl = {
			new ColumnHeader("担当者",		"担当者",    UseCsv._PERSON_NAME),
			new ColumnHeader("タスク数",    "数",        UseCsv._TASK_NUM_R),	//※右寄せ
			new ColumnHeader("Progress",    "平均進捗率",UseCsv._PROGRESS_BAR, true)
		};

		private readonly ColumnHeader[] NoLimitHeaderTbl = {
			new ColumnHeader("id",          "#",       UseCsv._TASK_ID),
			new ColumnHeader("題名",		"題名",    UseCsv._TASK_NAME),
			new ColumnHeader("担当者",		"担当者",  UseCsv._PERSON_NAME),
			new ColumnHeader("Progress",    "進捗率",  UseCsv._PROGRESS_BAR, true)
		};

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
							if (UseCsv.Tbl[column] == 0) continue;

							// 本体データの取得
							dataDic.Add(column, values[column]);

							#region 担当者別の進捗情報の取得

							if (column == UseCsv._PERSON_NAME) {
								string progressName = this.Nobody;
								if (values[UseCsv._PERSON_NAME] != "\"\"") {
									progressName = values[UseCsv._PERSON_NAME];
								}
								else {
									progressName = this.Nobody;
								}

								int rate = Convert.ToInt32(values[UseCsv._PROGRESS_RATE]);
								if (!personsTask.IsNewPerson(progressName, rate)) {
									personsTask.AddProgress(progressName, rate);
								}
							}

							#endregion 担当者別の進捗情報の取得

							if (column == UseCsv._PARENT) {
								this.ParentList.Add(values[UseCsv._PARENT]);
							}
						}

						rowDicList.Add(dataDic);
					}
				}
			}
			catch (IOException e) {
				MessageBox.Show("Excelで開いている場合は閉じてから使ってください\r\n" + e.Message,
					"エラー", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
			}
			catch (Exception e) {
				MessageBox.Show("例外が発生", "エラー\r\n" + e.Message, MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
			}
			MakeProgressRow(rowDicList);
			MakePersonTaskGrid(personsTask);
			MakeNoLimitTaskGrid();
		}

		private void InitDgvProgress()
		{
			this.DgvProgress.Init(this.groupBox3, "期日ありタスク進捗情報", ProgressHeaderTbl);
		}

		private void MakeProgressRow(List<Dictionary<int, string>> rowDicList)
		{
			this.DgvProgress.Rows.Clear();
			NoLimitList = new List<Dictionary<int, string>>();
			int dicRowCount = 0;

			//名前順にソート→ （親チケット）が先頭を先頭にする
			var sort = rowDicList
				.OrderBy(n => n[UseCsv._PERSON_NAME].ToString())
				.ThenBy(n => n[UseCsv._TASK_ID].ToString())
				.ToList();

			foreach (var dicCell in sort) {
				// 条件を満たしたタスクだけ表示
				if (dicCell[UseCsv._PROGRESS_RATE] == "100") continue;

				string name = dicCell[UseCsv._PERSON_NAME];
				string day = dicCell[UseCsv._DELIVERY_DAY];

				if (name == "\"\"") {
					dicCell[UseCsv._PERSON_NAME] = this.Nobody;
				}

				if (name == "\"\"" && day == "\"\"") {
					dicCell[UseCsv._PERSON_NAME] = this.Nobody;
					this.NoLimitList.Add(dicCell);
					continue;
				}
				else
				if (day == "\"\"") {
					this.NoLimitList.Add(dicCell);
					continue;
				}

				this.DgvProgress.Rows.Add();
				string myId = dicCell[UseCsv._TASK_ID];
				this.DgvProgress.Rows[dicRowCount].Cells["ID"].Value = myId;
				if (this.ParentList.Any(id => id == myId)) {
					this.DgvProgress.Rows[dicRowCount].Cells["ID"].Style.BackColor = Color.Yellow;
				}

				string parent = dicCell[UseCsv._PARENT];
				if (parent == "\"\"") {
					parent = String.Empty;
				}
				else {
					this.DgvProgress.Rows[dicRowCount].Cells["PARENT"].Style.ForeColor = Color.Red;
				}
				this.DgvProgress.Rows[dicRowCount].Cells["PARENT"].Value = parent;

				this.DgvProgress.Rows[dicRowCount].Cells["TITLE"].Value = dicCell[UseCsv._TASK_NAME];
				this.DgvProgress.Rows[dicRowCount].Cells["NAME"].Value = dicCell[UseCsv._PERSON_NAME];
				this.DgvProgress.Rows[dicRowCount].Cells["BAR"].Value = Convert.ToInt32(dicCell[UseCsv._PROGRESS_RATE]);
				this.DgvProgress.Rows[dicRowCount].Cells["LIMIT"].Value = dicCell[UseCsv._DELIVERY_DAY];

				DateTime dNow = DateTime.Now.Date;  // 時間なしの今日
				DateTime dTime = DateTime.Parse(dicCell[UseCsv._DELIVERY_DAY]);
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
			this.DgvMember.Init(this.groupBox2, "期日未定のタスク", MemberHeaderTbl);
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
			this.DgvNoLimitTask.Init(this.groupBox4, "期日未定のタスク", NoLimitHeaderTbl);
		}

		private void MakeNoLimitTaskGrid()
		{
			this.DgvNoLimitTask.Rows.Clear();

			int row = 0;

			//名前順にソート→ （親チケット）を先頭にする
			var sort = this.NoLimitList
				.OrderBy(n => n[UseCsv._PERSON_NAME].ToString())
				.ThenBy(n => n[UseCsv._TASK_ID].ToString())
				.ToList();

			foreach (var dic in sort) {
				this.DgvNoLimitTask.Rows.Add(dic[UseCsv._TASK_ID]);
				this.DgvNoLimitTask.Rows[row].Cells["題名"].Value = dic[UseCsv._TASK_NAME];
				this.DgvNoLimitTask.Rows[row].Cells["担当者"].Value = dic[UseCsv._PERSON_NAME];
				this.DgvNoLimitTask.Rows[row].Cells["Progress"].Value = Convert.ToInt32(dic[UseCsv._PROGRESS_RATE]);
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