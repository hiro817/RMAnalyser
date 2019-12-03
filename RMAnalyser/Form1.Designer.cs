namespace RMAnalyser
{
	partial class Form1
	{
		/// <summary>
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows フォーム デザイナーで生成されたコード

		/// <summary>
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.textBoxファイル名 = new System.Windows.Forms.TextBox();
			this.label情報 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.textBox開発 = new System.Windows.Forms.TextBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.button担当者別 = new System.Windows.Forms.Button();
			this.button期日あり進捗 = new System.Windows.Forms.Button();
			this.button期日未定タスク = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.textBoxファイル名);
			this.groupBox1.Location = new System.Drawing.Point(13, 13);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(251, 57);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "groupBox1";
			// 
			// textBoxファイル名
			// 
			this.textBoxファイル名.Location = new System.Drawing.Point(9, 19);
			this.textBoxファイル名.Name = "textBoxファイル名";
			this.textBoxファイル名.Size = new System.Drawing.Size(236, 19);
			this.textBoxファイル名.TabIndex = 1;
			// 
			// label情報
			// 
			this.label情報.AutoSize = true;
			this.label情報.ForeColor = System.Drawing.Color.Red;
			this.label情報.Location = new System.Drawing.Point(528, 334);
			this.label情報.Name = "label情報";
			this.label情報.Size = new System.Drawing.Size(35, 12);
			this.label情報.TabIndex = 0;
			this.label情報.Text = "label1";
			// 
			// groupBox2
			// 
			this.groupBox2.Location = new System.Drawing.Point(13, 77);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(251, 248);
			this.groupBox2.TabIndex = 1;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "groupBox2";
			// 
			// groupBox3
			// 
			this.groupBox3.Location = new System.Drawing.Point(271, 13);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(640, 312);
			this.groupBox3.TabIndex = 2;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "groupBox3";
			// 
			// textBox開発
			// 
			this.textBox開発.BackColor = System.Drawing.Color.Wheat;
			this.textBox開発.Location = new System.Drawing.Point(501, 358);
			this.textBox開発.Multiline = true;
			this.textBox開発.Name = "textBox開発";
			this.textBox開発.Size = new System.Drawing.Size(410, 235);
			this.textBox開発.TabIndex = 3;
			// 
			// groupBox4
			// 
			this.groupBox4.Location = new System.Drawing.Point(13, 358);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(482, 209);
			this.groupBox4.TabIndex = 4;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "groupBox4";
			// 
			// button担当者別
			// 
			this.button担当者別.Location = new System.Drawing.Point(13, 329);
			this.button担当者別.Name = "button担当者別";
			this.button担当者別.Size = new System.Drawing.Size(251, 23);
			this.button担当者別.TabIndex = 5;
			this.button担当者別.Text = "担当者別をコピー";
			this.button担当者別.UseVisualStyleBackColor = true;
			this.button担当者別.Click += new System.EventHandler(this.button担当者別_Click);
			// 
			// button期日あり進捗
			// 
			this.button期日あり進捗.Location = new System.Drawing.Point(271, 329);
			this.button期日あり進捗.Name = "button期日あり進捗";
			this.button期日あり進捗.Size = new System.Drawing.Size(251, 23);
			this.button期日あり進捗.TabIndex = 6;
			this.button期日あり進捗.Text = "期日あり進捗をコピー";
			this.button期日あり進捗.UseVisualStyleBackColor = true;
			this.button期日あり進捗.Click += new System.EventHandler(this.button期日あり進捗_Click);
			// 
			// button期日未定タスク
			// 
			this.button期日未定タスク.Location = new System.Drawing.Point(13, 570);
			this.button期日未定タスク.Name = "button期日未定タスク";
			this.button期日未定タスク.Size = new System.Drawing.Size(251, 23);
			this.button期日未定タスク.TabIndex = 7;
			this.button期日未定タスク.Text = "期日未定タスクをコピー";
			this.button期日未定タスク.UseVisualStyleBackColor = true;
			this.button期日未定タスク.Click += new System.EventHandler(this.button期日未定タスク_Click);
			// 
			// Form1
			// 
			this.AllowDrop = true;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(919, 605);
			this.Controls.Add(this.button期日未定タスク);
			this.Controls.Add(this.label情報);
			this.Controls.Add(this.button期日あり進捗);
			this.Controls.Add(this.button担当者別);
			this.Controls.Add(this.groupBox4);
			this.Controls.Add(this.textBox開発);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "Form1";
			this.Text = "Form1";
			this.DragDrop += new System.Windows.Forms.DragEventHandler(this.Form1_DragDrop);
			this.DragEnter += new System.Windows.Forms.DragEventHandler(this.Form1_DragEnter);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.Label label情報;
		private System.Windows.Forms.TextBox textBoxファイル名;
		private System.Windows.Forms.TextBox textBox開発;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Button button担当者別;
		private System.Windows.Forms.Button button期日あり進捗;
		private System.Windows.Forms.Button button期日未定タスク;
	}
}

