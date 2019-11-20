using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RMAnalyser
{
	class BaseDGV : DataGridView//※またはForm
	{
		DataGridView dgv;

		protected BaseDGV()
		{
			dgv = new DataGridView();

			// セルの編集不可
			dgv.ReadOnly = true;
			// 選択をセルだけに
			dgv.SelectionMode = DataGridViewSelectionMode.CellSelect;
			// 行を削除不可にする
			dgv.AllowUserToDeleteRows = false;
			// 列が自動で作成されないようにする
			dgv.AutoGenerateColumns = false;
			// 高さ・幅の変更不可にする
			dgv.AllowUserToResizeRows = false;
			dgv.AllowUserToResizeColumns = false;
			// 新しい行の追加不可に設定
			dgv.AllowUserToAddRows = false;
		}


	}

}
