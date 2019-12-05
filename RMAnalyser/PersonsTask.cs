using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RMAnalyser
{
	class PersonsTask
	{
		private Dictionary<string, List<float>> m_NameDic;

		public Dictionary<string, List<float>> NameDic
		{
			get { return this.m_NameDic; }
			private set { this.m_NameDic = value; }
		}

		public PersonsTask()
		{
			this.NameDic = new Dictionary<string, List<float>>();
		}

		public bool IsNewPerson(string name, float progress)
		{
			bool flg = NameDic.ContainsKey(name);
			// 初めての名前
			if (!flg && name != "")
			{
				var dataList = new List<float>();
				dataList.Add(progress);
				this.NameDic.Add(name, dataList);
				return true;
			}
			return false;
		}

		public void AddProgress(string name, float progress)
		{
			if (NameDic.ContainsKey(name))
			{
				this.NameDic[name].Add(progress);
			}
		}

		public float GetAverageProgress(string name) => m_NameDic[name].Average();

	}
}
