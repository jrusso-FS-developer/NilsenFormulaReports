namespace Races_CSV_To_Excel
{
    partial class frmRacesCSVtoExcel
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
            if (disposing && (components != null))
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmRacesCSVtoExcel));
            this.ofdRacesCSV = new System.Windows.Forms.OpenFileDialog();
            this.lblHeaderMessage = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.btnFileBrowse = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnProcess = new System.Windows.Forms.Button();
            this.txtConsole = new System.Windows.Forms.TextBox();
            this.lblDropDownMessage = new System.Windows.Forms.Label();
            this.ddReports = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // ofdRacesCSV
            // 
            this.ofdRacesCSV.FileName = string.Empty;
            this.ofdRacesCSV.FileOk += new System.ComponentModel.CancelEventHandler(this.ofdRacesCSV_FileOk);
            // 
            // lblHeaderMessage
            // 
            this.lblHeaderMessage.AutoSize = true;
            this.lblHeaderMessage.Location = new System.Drawing.Point(15, 73);
            this.lblHeaderMessage.Name = "lblHeaderMessage";
            this.lblHeaderMessage.Size = new System.Drawing.Size(249, 13);
            this.lblHeaderMessage.TabIndex = 0;
            this.lblHeaderMessage.Text = "Enter the path to your file, or click \'Browse\' to open.";
            // 
            // txtFileName
            // 
            this.txtFileName.CharacterCasing = System.Windows.Forms.CharacterCasing.Lower;
            this.txtFileName.Location = new System.Drawing.Point(15, 90);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(520, 20);
            this.txtFileName.TabIndex = 1;
            // 
            // btnFileBrowse
            // 
            this.btnFileBrowse.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnFileBrowse.Location = new System.Drawing.Point(541, 88);
            this.btnFileBrowse.Name = "btnFileBrowse";
            this.btnFileBrowse.Size = new System.Drawing.Size(97, 23);
            this.btnFileBrowse.TabIndex = 2;
            this.btnFileBrowse.Text = "Browse";
            this.btnFileBrowse.UseVisualStyleBackColor = true;
            this.btnFileBrowse.Click += new System.EventHandler(this.btnFileBrowse_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(500, 435);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(99, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnProcess
            // 
            this.btnProcess.Location = new System.Drawing.Point(605, 435);
            this.btnProcess.Name = "btnProcess";
            this.btnProcess.Size = new System.Drawing.Size(107, 23);
            this.btnProcess.TabIndex = 5;
            this.btnProcess.Text = "Process";
            this.btnProcess.UseVisualStyleBackColor = true;
            this.btnProcess.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // txtConsole
            // 
            this.txtConsole.BackColor = System.Drawing.SystemColors.MenuText;
            this.txtConsole.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtConsole.ForeColor = System.Drawing.Color.Lime;
            this.txtConsole.Location = new System.Drawing.Point(12, 117);
            this.txtConsole.Multiline = true;
            this.txtConsole.Name = "txtConsole";
            this.txtConsole.ReadOnly = true;
            this.txtConsole.Size = new System.Drawing.Size(700, 312);
            this.txtConsole.TabIndex = 0;
            // 
            // lblDropDownMessage
            // 
            this.lblDropDownMessage.AutoSize = true;
            this.lblDropDownMessage.Location = new System.Drawing.Point(15, 17);
            this.lblDropDownMessage.Name = "lblDropDownMessage";
            this.lblDropDownMessage.Size = new System.Drawing.Size(103, 13);
            this.lblDropDownMessage.TabIndex = 6;
            this.lblDropDownMessage.Text = "Choose Your Report";
            // 
            // ddReports
            // 
            this.ddReports.FormattingEnabled = true;
            this.ddReports.Items.AddRange(new object[] {
            "Turf Formula",
            "Pace Forecaster Formula"});
            this.ddReports.Location = new System.Drawing.Point(15, 33);
            this.ddReports.Name = "ddReports";
            this.ddReports.Size = new System.Drawing.Size(234, 21);
            this.ddReports.TabIndex = 7;
            // 
            // frmRacesCSVtoExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(726, 468);
            this.Controls.Add(this.ddReports);
            this.Controls.Add(this.lblDropDownMessage);
            this.Controls.Add(this.btnProcess);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.txtConsole);
            this.Controls.Add(this.btnFileBrowse);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.lblHeaderMessage);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmRacesCSVtoExcel";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Nilsen Race Formula Reports v 1.0";
            this.Load += new System.EventHandler(this.RacesCSVtoExcel_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog ofdRacesCSV;
        private System.Windows.Forms.Label lblHeaderMessage;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Button btnFileBrowse;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnProcess;
        private System.Windows.Forms.TextBox txtConsole;
        private System.Windows.Forms.Label lblDropDownMessage;
        private System.Windows.Forms.ComboBox ddReports;
    }
}

