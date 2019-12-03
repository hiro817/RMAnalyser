using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RMAnalyser
{
	class DGV : DataGridView
	{
		public DGV() : base()
		{
		}

		public void Init(GroupBox gbox, string titleName="")
		{
			gbox.Controls.Add(this);
			// 自動でコントロールの四辺にドッキングして適切なサイズに調整される
			this.Dock = DockStyle.Fill;

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
				// 左ヘッダをなくする(行番号を付けるイベントハンドラも削除すること)
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

			this.Location = new Point(10, 20);

		}

		public string ButtonClick()
		{
			this.SelectAll();
			this.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
			Clipboard.SetDataObject(this.GetClipboardContent());

			return "クリップボードにコピーしました！";
		}

	}
}
