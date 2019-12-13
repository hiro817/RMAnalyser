using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RMAnalyser
{
	static class UseCsv
	{
		public static int[] Tbl = {
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
			30,		// 22 タスク数
		};

		public const int _TASK_ID = 0;
		public const int _PARENT = 3;
		public const int _TASK_NAME = 6;
		public const int _PERSON_NAME = 8;
		public const int _DELIVERY_DAY = 13;
		public const int _PROGRESS_RATE = 15;
		public const int _REMAIMING = 20;
		public const int _PROGRESS_BAR = 21;
		public const int _TASK_NUM = 22;

	}
}
