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

		public void Init(GroupBox gbox)
		{
			gbox.Controls.Add(this);

			Font = new Font("Meiryo UI", 8);

			// セルの編集不可
			ReadOnly = true;

			// 選択をセルだけに
			SelectionMode = DataGridViewSelectionMode.CellSelect;

			// 行を削除不可にする
			AllowUserToDeleteRows = false;

			// 列が自動で作成されないようにする
			AutoGenerateColumns = false;

			// 高さ・幅の変更不可にする
			AllowUserToResizeRows = false;
			AllowUserToResizeColumns = false;

			// 新しい行の追加不可に設定
			AllowUserToAddRows = false;

			this.Location = new Point(10, 20);

		}
	}
}
