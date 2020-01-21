using System.Collections.Generic;
using System.Linq;

namespace GoogleApiIntegration
{
	public class ColHelper
	{
		public List<Col> Cols { get; set; }

		public ColHelper(int numCols)
		{
			Cols = new List<Col>();
			var charIndex = 0;
			for (var i = 0; i < numCols; i++)
			{
				Cols.Add(new Col
				{
					ColNum = i + 1,
					ColName = chars[i]
				});
				charIndex = charIndex == chars.Count ? 0 : charIndex + 1;
			}
		}

		public int ColToNum(string colName)
		{
			var col = Cols.FirstOrDefault(c => c.ColName == colName);
			return col?.ColNum ?? 1;
		}

		public string NumToCol(int colNum)
		{
			var col = Cols.FirstOrDefault(c => c.ColNum == colNum);
			return col?.ColName ?? "A";
		}

		private readonly List<string> chars = new List<string>
		{
			"A",
			"B",
			"C",
			"D",
			"E",
			"F",
			"G",
			"H",
			"I",
			"J",
			"K",
			"L",
			"M",
			"N",
			"O",
			"P",
			"Q",
			"R",
			"S",
			"T",
			"U",
			"V",
			"W",
			"X",
			"Y",
			"Z"
		};
	}

	public class Col
	{
		public string ColName { get; set; }
		public int ColNum { get; set; }
	}
}
