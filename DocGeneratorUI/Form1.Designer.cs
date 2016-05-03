namespace DocGenerator
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.btnSDDP = new System.Windows.Forms.Button();
			this.btnTest = new System.Windows.Forms.Button();
			this.buttonTestSpeed = new System.Windows.Forms.Button();
			this.Button_GenerateExcel = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnSDDP
			// 
			this.btnSDDP.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSDDP.ForeColor = System.Drawing.Color.MediumPurple;
			this.btnSDDP.Image = ((System.Drawing.Image)(resources.GetObject("btnSDDP.Image")));
			this.btnSDDP.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnSDDP.Location = new System.Drawing.Point(13, 12);
			this.btnSDDP.Name = "btnSDDP";
			this.btnSDDP.Size = new System.Drawing.Size(440, 43);
			this.btnSDDP.TabIndex = 3;
			this.btnSDDP.Text = "Process ACTIVE SharePoint Document Collections ...";
			this.btnSDDP.UseVisualStyleBackColor = true;
			this.btnSDDP.Click += new System.EventHandler(this.btnSDDP_Click);
			// 
			// btnTest
			// 
			this.btnTest.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.btnTest.ForeColor = System.Drawing.Color.RoyalBlue;
			this.btnTest.Image = ((System.Drawing.Image)(resources.GetObject("btnTest.Image")));
			this.btnTest.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.btnTest.Location = new System.Drawing.Point(12, 115);
			this.btnTest.Name = "btnTest";
			this.btnTest.Size = new System.Drawing.Size(440, 39);
			this.btnTest.TabIndex = 5;
			this.btnTest.Text = "Generate an MS Word Document";
			this.btnTest.UseVisualStyleBackColor = true;
			this.btnTest.Click += new System.EventHandler(this.btnOpenMSwordDocument);
			// 
			// buttonTestSpeed
			// 
			this.buttonTestSpeed.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.buttonTestSpeed.ForeColor = System.Drawing.Color.Goldenrod;
			this.buttonTestSpeed.Location = new System.Drawing.Point(12, 293);
			this.buttonTestSpeed.Name = "buttonTestSpeed";
			this.buttonTestSpeed.Size = new System.Drawing.Size(440, 39);
			this.buttonTestSpeed.TabIndex = 6;
			this.buttonTestSpeed.Text = "Data Access Speed Experiment";
			this.buttonTestSpeed.UseVisualStyleBackColor = true;
			this.buttonTestSpeed.Click += new System.EventHandler(this.buttonTestSpeed_Click);
			// 
			// Button_GenerateExcel
			// 
			this.Button_GenerateExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.Button_GenerateExcel.ForeColor = System.Drawing.Color.ForestGreen;
			this.Button_GenerateExcel.Image = ((System.Drawing.Image)(resources.GetObject("Button_GenerateExcel.Image")));
			this.Button_GenerateExcel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.Button_GenerateExcel.Location = new System.Drawing.Point(15, 160);
			this.Button_GenerateExcel.Name = "Button_GenerateExcel";
			this.Button_GenerateExcel.Size = new System.Drawing.Size(440, 39);
			this.Button_GenerateExcel.TabIndex = 7;
			this.Button_GenerateExcel.Text = "Generate an MS Excel Workbook";
			this.Button_GenerateExcel.UseVisualStyleBackColor = true;
			this.Button_GenerateExcel.Click += new System.EventHandler(this.Button_GenerateExcel_Click);
			// 
			// button1
			// 
			this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.button1.ForeColor = System.Drawing.Color.DarkGoldenrod;
			this.button1.Location = new System.Drawing.Point(12, 338);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(440, 39);
			this.button1.TabIndex = 8;
			this.button1.Text = "Test SQLite code Experiment";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.buttonSQLiteTest_Click);
			// 
			// Form1
			// 
			this.ClientSize = new System.Drawing.Size(467, 392);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.Button_GenerateExcel);
			this.Controls.Add(this.buttonTestSpeed);
			this.Controls.Add(this.btnTest);
			this.Controls.Add(this.btnSDDP);
			this.Name = "Form1";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

			}

		#endregion
		private System.Windows.Forms.Button btnSDDP;
		private System.Windows.Forms.Button btnTest;
		private System.Windows.Forms.Button buttonTestSpeed;
		private System.Windows.Forms.Button Button_GenerateExcel;
		private System.Windows.Forms.Button button1;
		}
	}

