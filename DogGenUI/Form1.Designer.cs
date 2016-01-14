namespace DogGenUI
	{
	partial class Form1
		{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
			{
			if(disposing && (components != null))
				{
				components.Dispose();
				}
			base.Dispose(disposing);
			}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
			{
			this.btnSDDP = new System.Windows.Forms.Button();
			this.lblConnect = new System.Windows.Forms.Label();
			this.btnTest = new System.Windows.Forms.Button();
			this.textBoxFileName = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btnSDDP
			// 
			this.btnSDDP.Location = new System.Drawing.Point(13, 12);
			this.btnSDDP.Name = "btnSDDP";
			this.btnSDDP.Size = new System.Drawing.Size(238, 23);
			this.btnSDDP.TabIndex = 3;
			this.btnSDDP.Text = "Get Document Collections to be Generated";
			this.btnSDDP.UseVisualStyleBackColor = true;
			this.btnSDDP.Click += new System.EventHandler(this.btnSDDP_Click);
			// 
			// lblConnect
			// 
			this.lblConnect.AutoSize = true;
			this.lblConnect.BackColor = System.Drawing.Color.DarkGray;
			this.lblConnect.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
			this.lblConnect.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.lblConnect.ForeColor = System.Drawing.Color.White;
			this.lblConnect.Location = new System.Drawing.Point(13, 38);
			this.lblConnect.MaximumSize = new System.Drawing.Size(500, 20);
			this.lblConnect.Name = "lblConnect";
			this.lblConnect.Size = new System.Drawing.Size(477, 15);
			this.lblConnect.TabIndex = 4;
			this.lblConnect.Text = "................................................................................." +
    "...........................................................................";
			this.lblConnect.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// btnTest
			// 
			this.btnTest.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F);
			this.btnTest.ForeColor = System.Drawing.Color.ForestGreen;
			this.btnTest.Location = new System.Drawing.Point(13, 210);
			this.btnTest.Name = "btnTest";
			this.btnTest.Size = new System.Drawing.Size(160, 23);
			this.btnTest.TabIndex = 5;
			this.btnTest.Text = "Open MS Word Document";
			this.btnTest.UseVisualStyleBackColor = true;
			this.btnTest.Click += new System.EventHandler(this.btnOpenMSwordDocument);
			// 
			// textBoxFileName
			// 
			this.textBoxFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
			this.textBoxFileName.Location = new System.Drawing.Point(181, 180);
			this.textBoxFileName.Multiline = true;
			this.textBoxFileName.Name = "textBoxFileName";
			this.textBoxFileName.Size = new System.Drawing.Size(732, 53);
			this.textBoxFileName.TabIndex = 6;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(10, 182);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(165, 15);
			this.label1.TabIndex = 7;
			this.label1.Text = "MS Word Document to open:";
			// 
			// Form1
			// 
			this.ClientSize = new System.Drawing.Size(925, 330);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.textBoxFileName);
			this.Controls.Add(this.btnTest);
			this.Controls.Add(this.lblConnect);
			this.Controls.Add(this.btnSDDP);
			this.Name = "Form1";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

			}

		#endregion
		private System.Windows.Forms.Button btnSDDP;
		private System.Windows.Forms.Label lblConnect;
		private System.Windows.Forms.Button btnTest;
		private System.Windows.Forms.TextBox textBoxFileName;
		private System.Windows.Forms.Label label1;
		}
	}

