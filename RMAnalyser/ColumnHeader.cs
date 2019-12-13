namespace RMAnalyser
{
	internal struct ColumnHeader
	{
		private int m_Index;

		public ColumnHeader(string id, string title, int tblIndex, bool isProgress = false)
		{
			this.Id = id;
			this.Title = title;
			this.m_Index = tblIndex;
			this.IsProgress = isProgress;
		}

		public string Id { get; private set; }
		public string Title { get; private set; }

		public int TblIndex		{ get { return this.m_Index; } }

		public bool IsProgress { get; set; }
	}
}