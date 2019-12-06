using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace RMAnalyser
{
	public static class DgvFile
	{
		public static void Out(DataGridView dgv, int addColumn = -1)
		{
			string filename = @"d:\test.csv";

			using (var writer = new StreamWriter(filename, false, Encoding.GetEncoding("shift_jis"))) {
				int colCount = dgv.Columns.Count;

				List<String> colList = new List<string>();
				for (int i = 0; i < colCount; i++) {
					colList.Add(dgv.Columns[i].HeaderCell.Value.ToString());
				}
				String[] headArray = colList.ToArray();
				// CSV形式に変換して出力
				writer.WriteLine(String.Join(",", headArray));

				// 行
				int rowCount = dgv.Rows.Count;
				for (int row = 0; row < rowCount; row++) {
					colList = new List<string>();
					// 列
					for (int col = 0; col < colCount; col++) {
						string data = dgv[col, row].Value.ToString();
						if (addColumn != -1 && addColumn == col) {
							data += "%";
						}
						colList.Add(data);
					}
					// 配列に変換
					String[] colArray = colList.ToArray();

					// CSV形式に変換
					String strCsvData = String.Join(",", colArray);
					writer.WriteLine(strCsvData);
				}
			}
		}
	}
}