using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RMAnalyser
{
	public static class ReadCSV
	{
		public static List<string[]> ReadFile(string filepath)
		{
			var lines = File.ReadLines(filepath);
			return lines
				.Skip(1)    // １行目はカラム名
				.Select(line => line.Split(','))
				.ToList();
		}

	}

}