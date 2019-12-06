using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace RMAnalyser
{
	class DGV : DataGridView
	{
		public DGV() : base()
		{
		}

		public GroupBox ParentGroupBox { get; set; }
		public string GBText { get; set; }

		public void Init(GroupBox groupBox, string gbText, string titleName = "")
		{
			this.GBText = gbText;

			this.ParentGroupBox = groupBox;
			groupBox.Controls.Add(this);

			// 自動でコントロールの四辺にドッキングして適切なサイズに調整される
			this.Dock = DockStyle.Fill;
			this.Location = new Point(10, 20);

			// 左上にタイトルがあるときだけ行番号をつける
			if (titleName != "") {
				// 左上隅のセルの値
				this.TopLeftHeaderCell.Value = titleName;

				this.RowPostPaint += delegate (object sender, DataGridViewRowPostPaintEventArgs e)
				{
					// 行ヘッダのセル領域を、行番号を描画する長方形とする
					// （ただし右端に4ドットのすき間を空ける）
					Rectangle rect = new Rectangle(
					  e.RowBounds.Location.X,
					  e.RowBounds.Location.Y,
					  this.RowHeadersWidth - 4,
					  e.RowBounds.Height);

					// 上記の長方形内に行番号を縦方向中央＆右詰めで描画する
					// フォントや前景色は行ヘッダの既定値を使用する
					TextRenderer.DrawText(
					  e.Graphics,
					  (e.RowIndex + 1).ToString(),
					  this.RowHeadersDefaultCellStyle.Font,
					  rect,
					  this.RowHeadersDefaultCellStyle.ForeColor,
					  TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
				};
			}
			else {
				// 左ヘッダをなくする
				this.RowHeadersVisible = false;
			}

			this.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

			this.Font = new Font("Meiryo UI", 8);

			// セルの編集不可
			this.ReadOnly = true;

			// 選択をセルだけに
			this.SelectionMode = DataGridViewSelectionMode.CellSelect;

			// 行を削除不可にする
			this.AllowUserToDeleteRows = false;

			// 列が自動で作成されないようにする
			this.AutoGenerateColumns = false;

			// 高さ・幅の変更不可にする
			this.AllowUserToResizeRows = false;
			this.AllowUserToResizeColumns = false;

			// 新しい行の追加不可に設定
			this.AllowUserToAddRows = false;

			this.ScrollBars = ScrollBars.Vertical;
			//VScroll = true;

			SetGroupTextRowCount();
		}

#if false//未使用になる予定
		public string ButtonClick()
		{
			this.SelectAll();
			this.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
			Clipboard.SetDataObject(this.GetClipboardContent());

			return "クリップボードにコピーしました！";
		}
#endif
		public void CopyToClipboard(int percentColumn = -1)
		{
			string copyText = "";


			int colCount = this.Columns.Count;

			//List<String> colList = new List<string>();
			for (int i = 0; i < colCount; i++) {
				//colList.Add(this.Columns[i].HeaderCell.Value.ToString());

				copyText += this.Columns[i].HeaderCell.Value.ToString() + "\t";
			}
			copyText += "\r\n";

			//String[] headArray = colList.ToArray();
			//String headData = String.Join("\t", headArray);


			// 行
			int rowCount = this.Rows.Count;
			for (int row = 0; row < rowCount; row++) {

				// 列
				for (int col = 0; col < colCount; col++) {
					string data = this[col, row].Value.ToString();
					if (percentColumn != -1 && percentColumn == col) {
						data += "%";
					}
					//colList.Add(data + "\t");
					copyText += data + "\t";
				}
				// 配列に変換
				//String[] colArray = colList.ToArray();
				copyText += "\r\n";

				// CSV形式に変換
				//String strCsvData = String.Join(",", colArray);
				//writer.WriteLine(strCsvData);
			}

			//String[] colArray = colList.ToArray();

			//String strCsvData = String.Join("\t", colArray);

			this.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
			//Clipboard.SetData(DataFormats.Text, colArray);
			Clipboard.SetData(DataFormats.Text, copyText);

		}

		public void SetGroupTextRowCount()
		{
			if (this.GBText != null) {
				this.ParentGroupBox.Text = this.GBText + " (" + this.RowCount.ToString() + ")";
			}
		}

	}
}
