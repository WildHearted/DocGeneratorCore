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
			this.btnButton2 = new System.Windows.Forms.Button();
			this.btnSDDP = new System.Windows.Forms.Button();
			this.lblConnect = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btnButton2
			// 
			this.btnButton2.Location = new System.Drawing.Point(692, 12);
			this.btnButton2.Name = "btnButton2";
			this.btnButton2.Size = new System.Drawing.Size(105, 23);
			this.btnButton2.TabIndex = 2;
			this.btnButton2.Text = "Comic action";
			this.btnButton2.UseVisualStyleBackColor = true;
			this.btnButton2.Visible = false;
			this.btnButton2.Click += new System.EventHandler(this.btnButton2_Click);
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
			// Form1
			// 
			this.ClientSize = new System.Drawing.Size(809, 330);
			this.Controls.Add(this.lblConnect);
			this.Controls.Add(this.btnSDDP);
			this.Controls.Add(this.btnButton2);
			this.Name = "Form1";
			this.ResumeLayout(false);
			this.PerformLayout();

			}

		#endregion
		private System.Windows.Forms.Button btnButton2;
		private System.Windows.Forms.Button btnSDDP;
		private System.Windows.Forms.Label lblConnect;
		}
	}

