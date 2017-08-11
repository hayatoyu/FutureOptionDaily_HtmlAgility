namespace FutureOptionDaily_HtmlAgility
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnSelectFilePath = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.dateTimePicker2 = new System.Windows.Forms.DateTimePicker();
            this.btnStart = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            this.ckBoxDuration = new System.Windows.Forms.CheckBox();
            this.rtbSelectedDate = new System.Windows.Forms.RichTextBox();
            this.btnResetDuration = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txtSaveDir = new System.Windows.Forms.TextBox();
            this.btn_SelectDir = new System.Windows.Forms.Button();
            this.btn_Split = new System.Windows.Forms.Button();
            this.btn_Duration = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(69, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Excel路徑：";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(138, 34);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(310, 22);
            this.txtFilePath.TabIndex = 1;
            // 
            // btnSelectFilePath
            // 
            this.btnSelectFilePath.Location = new System.Drawing.Point(454, 32);
            this.btnSelectFilePath.Name = "btnSelectFilePath";
            this.btnSelectFilePath.Size = new System.Drawing.Size(25, 23);
            this.btnSelectFilePath.TabIndex = 2;
            this.btnSelectFilePath.Text = "...";
            this.btnSelectFilePath.UseVisualStyleBackColor = true;
            this.btnSelectFilePath.Click += new System.EventHandler(this.btnSelectFilePath_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(71, 147);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 3;
            this.label2.Text = "本交易日：";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(59, 186);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "上一交易日：";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(138, 143);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(200, 22);
            this.dateTimePicker1.TabIndex = 5;
            // 
            // dateTimePicker2
            // 
            this.dateTimePicker2.Location = new System.Drawing.Point(138, 183);
            this.dateTimePicker2.Name = "dateTimePicker2";
            this.dateTimePicker2.Size = new System.Drawing.Size(200, 22);
            this.dateTimePicker2.TabIndex = 6;
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(719, 382);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(75, 23);
            this.btnStart.TabIndex = 7;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoEllipsis = true;
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(69, 230);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 12);
            this.lblStatus.TabIndex = 8;
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.Location = new System.Drawing.Point(574, 34);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 9;
            this.monthCalendar1.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateSelected);
            // 
            // ckBoxDuration
            // 
            this.ckBoxDuration.AutoSize = true;
            this.ckBoxDuration.Location = new System.Drawing.Point(574, 12);
            this.ckBoxDuration.Name = "ckBoxDuration";
            this.ckBoxDuration.Size = new System.Drawing.Size(72, 16);
            this.ckBoxDuration.TabIndex = 10;
            this.ckBoxDuration.Text = "抓取期間";
            this.ckBoxDuration.UseVisualStyleBackColor = true;
            // 
            // rtbSelectedDate
            // 
            this.rtbSelectedDate.Location = new System.Drawing.Point(574, 236);
            this.rtbSelectedDate.Name = "rtbSelectedDate";
            this.rtbSelectedDate.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.rtbSelectedDate.Size = new System.Drawing.Size(220, 140);
            this.rtbSelectedDate.TabIndex = 11;
            this.rtbSelectedDate.Text = "";
            // 
            // btnResetDuration
            // 
            this.btnResetDuration.Location = new System.Drawing.Point(719, 207);
            this.btnResetDuration.Name = "btnResetDuration";
            this.btnResetDuration.Size = new System.Drawing.Size(75, 23);
            this.btnResetDuration.TabIndex = 12;
            this.btnResetDuration.Text = "重選期間";
            this.btnResetDuration.UseVisualStyleBackColor = true;
            this.btnResetDuration.Click += new System.EventHandler(this.btnResetDuration_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(71, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 13;
            this.label4.Text = "拆分路徑：";
            // 
            // txtSaveDir
            // 
            this.txtSaveDir.Location = new System.Drawing.Point(138, 88);
            this.txtSaveDir.Name = "txtSaveDir";
            this.txtSaveDir.Size = new System.Drawing.Size(310, 22);
            this.txtSaveDir.TabIndex = 14;
            // 
            // btn_SelectDir
            // 
            this.btn_SelectDir.Location = new System.Drawing.Point(455, 87);
            this.btn_SelectDir.Name = "btn_SelectDir";
            this.btn_SelectDir.Size = new System.Drawing.Size(24, 23);
            this.btn_SelectDir.TabIndex = 15;
            this.btn_SelectDir.Text = "...";
            this.btn_SelectDir.UseVisualStyleBackColor = true;
            this.btn_SelectDir.Click += new System.EventHandler(this.btn_SelectDir_Click);
            // 
            // btn_Split
            // 
            this.btn_Split.Location = new System.Drawing.Point(574, 382);
            this.btn_Split.Name = "btn_Split";
            this.btn_Split.Size = new System.Drawing.Size(75, 23);
            this.btn_Split.TabIndex = 16;
            this.btn_Split.Text = "拆分Excel";
            this.btn_Split.UseVisualStyleBackColor = true;
            this.btn_Split.Click += new System.EventHandler(this.btn_Split_Click);
            // 
            // btn_Duration
            // 
            this.btn_Duration.Location = new System.Drawing.Point(719, 7);
            this.btn_Duration.Name = "btn_Duration";
            this.btn_Duration.Size = new System.Drawing.Size(75, 23);
            this.btn_Duration.TabIndex = 17;
            this.btn_Duration.Text = "輸入期間";
            this.btn_Duration.UseVisualStyleBackColor = true;
            this.btn_Duration.Click += new System.EventHandler(this.btn_Duration_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(938, 417);
            this.Controls.Add(this.btn_Duration);
            this.Controls.Add(this.btn_Split);
            this.Controls.Add(this.btn_SelectDir);
            this.Controls.Add(this.txtSaveDir);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnResetDuration);
            this.Controls.Add(this.rtbSelectedDate);
            this.Controls.Add(this.ckBoxDuration);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.dateTimePicker2);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnSelectFilePath);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.label1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnSelectFilePath;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.DateTimePicker dateTimePicker2;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
        private System.Windows.Forms.CheckBox ckBoxDuration;
        public System.Windows.Forms.RichTextBox rtbSelectedDate;
        private System.Windows.Forms.Button btnResetDuration;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtSaveDir;
        private System.Windows.Forms.Button btn_SelectDir;
        private System.Windows.Forms.Button btn_Split;
        private System.Windows.Forms.Button btn_Duration;
    }
}

