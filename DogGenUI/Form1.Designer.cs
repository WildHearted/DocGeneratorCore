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
			this.btnSDDP = new System.Windows.Forms.Button();
			this.btnTest = new System.Windows.Forms.Button();
			this.buttonTestSpeed = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnSDDP
			// 
			this.btnSDDP.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.btnSDDP.Location = new System.Drawing.Point(13, 12);
			this.btnSDDP.Name = "btnSDDP";
			this.btnSDDP.Size = new System.Drawing.Size(440, 43);
			this.btnSDDP.TabIndex = 3;
			this.btnSDDP.Text = "Process all active Document Collections currently active in SharePoint...";
			this.btnSDDP.UseVisualStyleBackColor = true;
			this.btnSDDP.Click += new System.EventHandler(this.btnSDDP_Click);
			// 
			// btnTest
			// 
			this.btnTest.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.btnTest.ForeColor = System.Drawing.Color.ForestGreen;
			this.btnTest.Location = new System.Drawing.Point(13, 61);
			this.btnTest.Name = "btnTest";
			this.btnTest.Size = new System.Drawing.Size(440, 39);
			this.btnTest.TabIndex = 5;
			this.btnTest.Text = "Process HTML text file located on local computer...";
			this.btnTest.UseVisualStyleBackColor = true;
			this.btnTest.Click += new System.EventHandler(this.btnOpenMSwordDocument);
			// 
			// buttonTestSpeed
			// 
			this.buttonTestSpeed.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
			this.buttonTestSpeed.ForeColor = System.Drawing.Color.ForestGreen;
			this.buttonTestSpeed.Location = new System.Drawing.Point(12, 106);
			this.buttonTestSpeed.Name = "buttonTestSpeed";
			this.buttonTestSpeed.Size = new System.Drawing.Size(440, 39);
			this.buttonTestSpeed.TabIndex = 6;
			this.buttonTestSpeed.Text = "Test Faster Data Access";
			this.buttonTestSpeed.UseVisualStyleBackColor = true;
			this.buttonTestSpeed.Click += new System.EventHandler(this.buttonTestSpeed_Click);
			// 
			// Form1
			// 
			this.ClientSize = new System.Drawing.Size(467, 330);
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
		}
	}

