namespace AESTest2._0
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.labelTime = new System.Windows.Forms.Label();
            this.lblQuestionText = new System.Windows.Forms.Label();
            this.lblAnswerD = new System.Windows.Forms.Label();
            this.lblAnswerC = new System.Windows.Forms.Label();
            this.lblAnswerB = new System.Windows.Forms.Label();
            this.lblAnswerA = new System.Windows.Forms.Label();
            this.lblQuestion = new System.Windows.Forms.Label();
            this.btnPrev = new System.Windows.Forms.Button();
            this.pBar = new System.Windows.Forms.ProgressBar();
            this.btnEnd = new System.Windows.Forms.Button();
            this.stage_2 = new System.Windows.Forms.Panel();
            this.btnNext = new System.Windows.Forms.Button();
            this.Time = new System.Windows.Forms.Timer(this.components);
            this.stage_1 = new System.Windows.Forms.Panel();
            this.btnStart = new System.Windows.Forms.Button();
            this.cmbPosts = new System.Windows.Forms.ComboBox();
            this.lblPosts = new System.Windows.Forms.Label();
            this.lblGroups = new System.Windows.Forms.Label();
            this.lblNames = new System.Windows.Forms.Label();
            this.cmbGroups = new System.Windows.Forms.ComboBox();
            this.cmbNames = new System.Windows.Forms.ComboBox();
            this.stage_3 = new System.Windows.Forms.Panel();
            this.btnNextTest = new System.Windows.Forms.Button();
            this.lblMark = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.YesNoLabel = new System.Windows.Forms.TextBox();
            this.stage_2.SuspendLayout();
            this.stage_1.SuspendLayout();
            this.stage_3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // labelTime
            // 
            this.labelTime.AutoSize = true;
            this.labelTime.BackColor = System.Drawing.Color.White;
            this.labelTime.Font = new System.Drawing.Font("Microsoft Sans Serif", 12.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelTime.ForeColor = System.Drawing.SystemColors.WindowText;
            this.labelTime.Location = new System.Drawing.Point(518, 0);
            this.labelTime.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelTime.Name = "labelTime";
            this.labelTime.Size = new System.Drawing.Size(247, 25);
            this.labelTime.TabIndex = 37;
            this.labelTime.Text = "Оставащо време: 30:00";
            this.labelTime.Click += new System.EventHandler(this.labelTime_Click);
            // 
            // lblQuestionText
            // 
            this.lblQuestionText.AutoSize = true;
            this.lblQuestionText.BackColor = System.Drawing.Color.White;
            this.lblQuestionText.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQuestionText.Font = new System.Drawing.Font("Microsoft Sans Serif", 22.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQuestionText.Location = new System.Drawing.Point(166, 23);
            this.lblQuestionText.Name = "lblQuestionText";
            this.lblQuestionText.Size = new System.Drawing.Size(55, 46);
            this.lblQuestionText.TabIndex = 51;
            this.lblQuestionText.Text = "...";
            // 
            // lblAnswerD
            // 
            this.lblAnswerD.AutoSize = true;
            this.lblAnswerD.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAnswerD.Location = new System.Drawing.Point(940, 437);
            this.lblAnswerD.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAnswerD.MaximumSize = new System.Drawing.Size(667, 0);
            this.lblAnswerD.Name = "lblAnswerD";
            this.lblAnswerD.Size = new System.Drawing.Size(147, 32);
            this.lblAnswerD.TabIndex = 44;
            this.lblAnswerD.Text = "Отговор Г";
            this.lblAnswerD.Click += new System.EventHandler(this.lblAnswerD_Click);
            // 
            // lblAnswerC
            // 
            this.lblAnswerC.AutoSize = true;
            this.lblAnswerC.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAnswerC.Location = new System.Drawing.Point(99, 437);
            this.lblAnswerC.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAnswerC.MaximumSize = new System.Drawing.Size(667, 0);
            this.lblAnswerC.Name = "lblAnswerC";
            this.lblAnswerC.Size = new System.Drawing.Size(150, 32);
            this.lblAnswerC.TabIndex = 43;
            this.lblAnswerC.Text = "Отговор В";
            this.lblAnswerC.Click += new System.EventHandler(this.lblAnswerC_Click);
            // 
            // lblAnswerB
            // 
            this.lblAnswerB.AutoSize = true;
            this.lblAnswerB.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAnswerB.Location = new System.Drawing.Point(940, 197);
            this.lblAnswerB.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAnswerB.MaximumSize = new System.Drawing.Size(667, 0);
            this.lblAnswerB.Name = "lblAnswerB";
            this.lblAnswerB.Size = new System.Drawing.Size(149, 32);
            this.lblAnswerB.TabIndex = 42;
            this.lblAnswerB.Text = "Отговор Б";
            this.lblAnswerB.Click += new System.EventHandler(this.lblAnswerB_Click);
            // 
            // lblAnswerA
            // 
            this.lblAnswerA.AutoSize = true;
            this.lblAnswerA.BackColor = System.Drawing.Color.Transparent;
            this.lblAnswerA.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAnswerA.Location = new System.Drawing.Point(99, 197);
            this.lblAnswerA.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblAnswerA.MaximumSize = new System.Drawing.Size(667, 0);
            this.lblAnswerA.Name = "lblAnswerA";
            this.lblAnswerA.Size = new System.Drawing.Size(150, 32);
            this.lblAnswerA.TabIndex = 41;
            this.lblAnswerA.Text = "Отговор А";
            this.lblAnswerA.Click += new System.EventHandler(this.lblAnswerA_Click);
            // 
            // lblQuestion
            // 
            this.lblQuestion.AutoSize = true;
            this.lblQuestion.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblQuestion.Location = new System.Drawing.Point(15, 28);
            this.lblQuestion.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblQuestion.Name = "lblQuestion";
            this.lblQuestion.Size = new System.Drawing.Size(144, 39);
            this.lblQuestion.TabIndex = 40;
            this.lblQuestion.Text = "Въпрос:";
            // 
            // btnPrev
            // 
            this.btnPrev.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnPrev.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnPrev.Font = new System.Drawing.Font("Microsoft Sans Serif", 18.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPrev.ForeColor = System.Drawing.Color.White;
            this.btnPrev.Location = new System.Drawing.Point(15, 644);
            this.btnPrev.Margin = new System.Windows.Forms.Padding(4);
            this.btnPrev.Name = "btnPrev";
            this.btnPrev.Size = new System.Drawing.Size(351, 67);
            this.btnPrev.TabIndex = 45;
            this.btnPrev.Text = "Предишен Въпрос";
            this.btnPrev.UseVisualStyleBackColor = false;
            this.btnPrev.Click += new System.EventHandler(this.btnPrev_Click);
            // 
            // pBar
            // 
            this.pBar.ForeColor = System.Drawing.Color.DodgerBlue;
            this.pBar.Location = new System.Drawing.Point(374, 644);
            this.pBar.Margin = new System.Windows.Forms.Padding(4);
            this.pBar.Name = "pBar";
            this.pBar.Size = new System.Drawing.Size(557, 28);
            this.pBar.Step = 14;
            this.pBar.TabIndex = 49;
            this.pBar.Visible = false;
            // 
            // btnEnd
            // 
            this.btnEnd.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnEnd.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnEnd.Font = new System.Drawing.Font("Microsoft Sans Serif", 18.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEnd.ForeColor = System.Drawing.Color.White;
            this.btnEnd.Location = new System.Drawing.Point(477, 585);
            this.btnEnd.Margin = new System.Windows.Forms.Padding(4);
            this.btnEnd.Name = "btnEnd";
            this.btnEnd.Size = new System.Drawing.Size(332, 67);
            this.btnEnd.TabIndex = 50;
            this.btnEnd.Text = "Приключи Теста";
            this.btnEnd.UseVisualStyleBackColor = false;
            this.btnEnd.Visible = false;
            this.btnEnd.Click += new System.EventHandler(this.btnEnd_Click);
            // 
            // stage_2
            // 
            this.stage_2.BackColor = System.Drawing.Color.White;
            this.stage_2.Controls.Add(this.lblQuestionText);
            this.stage_2.Controls.Add(this.btnNext);
            this.stage_2.Controls.Add(this.lblQuestion);
            this.stage_2.Controls.Add(this.btnEnd);
            this.stage_2.Controls.Add(this.pBar);
            this.stage_2.Controls.Add(this.btnPrev);
            this.stage_2.Controls.Add(this.lblAnswerD);
            this.stage_2.Controls.Add(this.lblAnswerC);
            this.stage_2.Controls.Add(this.lblAnswerB);
            this.stage_2.Controls.Add(this.lblAnswerA);
            this.stage_2.Location = new System.Drawing.Point(12, 12);
            this.stage_2.Name = "stage_2";
            this.stage_2.Size = new System.Drawing.Size(934, 738);
            this.stage_2.TabIndex = 51;
            // 
            // btnNext
            // 
            this.btnNext.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnNext.Font = new System.Drawing.Font("Microsoft Sans Serif", 18.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNext.ForeColor = System.Drawing.Color.White;
            this.btnNext.Location = new System.Drawing.Point(939, 644);
            this.btnNext.Margin = new System.Windows.Forms.Padding(4);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(329, 67);
            this.btnNext.TabIndex = 51;
            this.btnNext.Text = "Следващ Въпрос";
            this.btnNext.UseVisualStyleBackColor = false;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // Time
            // 
            this.Time.Interval = 1000;
            this.Time.Tick += new System.EventHandler(this.Time_Tick);
            // 
            // stage_1
            // 
            this.stage_1.BackColor = System.Drawing.Color.White;
            this.stage_1.BackgroundImage = global::AESTest2._0.Properties.Resources.tpp_aes_logo;
            this.stage_1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.stage_1.Controls.Add(this.btnStart);
            this.stage_1.Controls.Add(this.cmbPosts);
            this.stage_1.Controls.Add(this.lblPosts);
            this.stage_1.Controls.Add(this.lblGroups);
            this.stage_1.Controls.Add(this.lblNames);
            this.stage_1.Controls.Add(this.cmbGroups);
            this.stage_1.Controls.Add(this.cmbNames);
            this.stage_1.Location = new System.Drawing.Point(143, 37);
            this.stage_1.Name = "stage_1";
            this.stage_1.Size = new System.Drawing.Size(1008, 226);
            this.stage_1.TabIndex = 38;
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnStart.FlatAppearance.BorderSize = 0;
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnStart.Font = new System.Drawing.Font("Microsoft Sans Serif", 18.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnStart.ForeColor = System.Drawing.Color.White;
            this.btnStart.Location = new System.Drawing.Point(10, 152);
            this.btnStart.Margin = new System.Windows.Forms.Padding(4);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(332, 67);
            this.btnStart.TabIndex = 48;
            this.btnStart.Text = "Започни Теста";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Visible = false;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // cmbPosts
            // 
            this.cmbPosts.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbPosts.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbPosts.FormattingEnabled = true;
            this.cmbPosts.Location = new System.Drawing.Point(169, 110);
            this.cmbPosts.Margin = new System.Windows.Forms.Padding(4);
            this.cmbPosts.Name = "cmbPosts";
            this.cmbPosts.Size = new System.Drawing.Size(773, 34);
            this.cmbPosts.TabIndex = 39;
            this.cmbPosts.SelectedIndexChanged += new System.EventHandler(this.cmb_SelectedIndexChanged);
            // 
            // lblPosts
            // 
            this.lblPosts.AutoSize = true;
            this.lblPosts.BackColor = System.Drawing.Color.White;
            this.lblPosts.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPosts.Location = new System.Drawing.Point(4, 110);
            this.lblPosts.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPosts.Name = "lblPosts";
            this.lblPosts.Size = new System.Drawing.Size(157, 32);
            this.lblPosts.TabIndex = 38;
            this.lblPosts.Text = "Длъжност:";
            // 
            // lblGroups
            // 
            this.lblGroups.AutoSize = true;
            this.lblGroups.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblGroups.Location = new System.Drawing.Point(4, 70);
            this.lblGroups.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblGroups.Name = "lblGroups";
            this.lblGroups.Size = new System.Drawing.Size(102, 32);
            this.lblGroups.TabIndex = 37;
            this.lblGroups.Text = "Изпит:";
            // 
            // lblNames
            // 
            this.lblNames.AutoSize = true;
            this.lblNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblNames.Location = new System.Drawing.Point(4, 28);
            this.lblNames.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblNames.Name = "lblNames";
            this.lblNames.Size = new System.Drawing.Size(79, 32);
            this.lblNames.TabIndex = 36;
            this.lblNames.Text = "Име:";
            // 
            // cmbGroups
            // 
            this.cmbGroups.AutoCompleteCustomSource.AddRange(new string[] {
            "изпит за квалификационна група по ПБЗРЕУ  група 2",
            "изпит за квалификационна група по ПБЗРЕУ  група 3",
            "изпит за квалификационна група по ПБЗРЕУ  група 4",
            "изпит за квалификационна група по ПБЗРЕУ  група 5",
            "изпит за квалификационна група по ПБЗРНЕУ  група 2",
            "изпит за квалификационна група по ПБЗРНЕУ  група 3",
            "изпит за квалификационна група по ПБЗРНЕУ  група 4",
            "изпит за квалификационна група по ПБЗРНЕУ  група 5",
            "изпит по наредба 9"});
            this.cmbGroups.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbGroups.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbGroups.FormattingEnabled = true;
            this.cmbGroups.Items.AddRange(new object[] {
            "изпит за квалификационна група по ПБЗРЕУ  група 2",
            "изпит за квалификационна група по ПБЗРЕУ  група 3",
            "изпит за квалификационна група по ПБЗРЕУ  група 4",
            "изпит за квалификационна група по ПБЗРЕУ  група 5",
            "изпит за квалификационна група по ПБЗРНЕУ  група 2",
            "изпит за квалификационна група по ПБЗРНЕУ  група 3",
            "изпит за квалификационна група по ПБЗРНЕУ  група 4",
            "изпит за квалификационна група по ПБЗРНЕУ  група 5",
            "изпит по наредба 9"});
            this.cmbGroups.Location = new System.Drawing.Point(169, 68);
            this.cmbGroups.Margin = new System.Windows.Forms.Padding(4);
            this.cmbGroups.Name = "cmbGroups";
            this.cmbGroups.Size = new System.Drawing.Size(773, 34);
            this.cmbGroups.TabIndex = 35;
            this.cmbGroups.SelectedIndexChanged += new System.EventHandler(this.cmb_SelectedIndexChanged);
            // 
            // cmbNames
            // 
            this.cmbNames.Cursor = System.Windows.Forms.Cursors.Default;
            this.cmbNames.DropDownHeight = 500;
            this.cmbNames.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbNames.FormattingEnabled = true;
            this.cmbNames.IntegralHeight = false;
            this.cmbNames.Location = new System.Drawing.Point(169, 21);
            this.cmbNames.Margin = new System.Windows.Forms.Padding(4);
            this.cmbNames.MaxDropDownItems = 20;
            this.cmbNames.Name = "cmbNames";
            this.cmbNames.Size = new System.Drawing.Size(773, 39);
            this.cmbNames.TabIndex = 34;
            this.cmbNames.SelectedIndexChanged += new System.EventHandler(this.cmb_SelectedIndexChanged);
            // 
            // stage_3
            // 
            this.stage_3.BackColor = System.Drawing.Color.White;
            this.stage_3.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.stage_3.Controls.Add(this.YesNoLabel);
            this.stage_3.Controls.Add(this.btnNextTest);
            this.stage_3.Controls.Add(this.lblMark);
            this.stage_3.Controls.Add(this.pictureBox1);
            this.stage_3.Enabled = false;
            this.stage_3.Location = new System.Drawing.Point(958, 269);
            this.stage_3.Name = "stage_3";
            this.stage_3.Size = new System.Drawing.Size(352, 380);
            this.stage_3.TabIndex = 36;
            this.stage_3.Visible = false;
            // 
            // btnNextTest
            // 
            this.btnNextTest.BackColor = System.Drawing.Color.DodgerBlue;
            this.btnNextTest.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnNextTest.Font = new System.Drawing.Font("Microsoft Sans Serif", 16.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNextTest.ForeColor = System.Drawing.Color.White;
            this.btnNextTest.Location = new System.Drawing.Point(34, 323);
            this.btnNextTest.Name = "btnNextTest";
            this.btnNextTest.Size = new System.Drawing.Size(288, 54);
            this.btnNextTest.TabIndex = 2;
            this.btnNextTest.Text = "Нов изпит";
            this.btnNextTest.UseVisualStyleBackColor = false;
            this.btnNextTest.Click += new System.EventHandler(this.btnNextTest_Click);
            // 
            // lblMark
            // 
            this.lblMark.AutoSize = true;
            this.lblMark.Location = new System.Drawing.Point(84, 0);
            this.lblMark.Name = "lblMark";
            this.lblMark.Size = new System.Drawing.Size(190, 17);
            this.lblMark.TabIndex = 1;
            this.lblMark.Text = "Оценка от последен изпит:";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImage = global::AESTest2._0.Properties.Resources.aeslogo;
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.pictureBox1.Location = new System.Drawing.Point(3, 7);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(346, 219);
            this.pictureBox1.TabIndex = 3;
            this.pictureBox1.TabStop = false;
            // 
            // YesNoLabel
            // 
            this.YesNoLabel.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.YesNoLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 48F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.YesNoLabel.Location = new System.Drawing.Point(34, 226);
            this.YesNoLabel.Name = "YesNoLabel";
            this.YesNoLabel.Size = new System.Drawing.Size(288, 91);
            this.YesNoLabel.TabIndex = 4;
            this.YesNoLabel.Text = "lblYesNo";
            this.YesNoLabel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.YesNoLabel.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1322, 790);
            this.ControlBox = false;
            this.Controls.Add(this.stage_2);
            this.Controls.Add(this.stage_1);
            this.Controls.Add(this.labelTime);
            this.Controls.Add(this.stage_3);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.Text = "AESTest 2.0";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.stage_2.ResumeLayout(false);
            this.stage_2.PerformLayout();
            this.stage_1.ResumeLayout(false);
            this.stage_1.PerformLayout();
            this.stage_3.ResumeLayout(false);
            this.stage_3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel stage_3;
        private System.Windows.Forms.Button btnNextTest;
        private System.Windows.Forms.Label lblMark;
        private System.Windows.Forms.Label labelTime;
        private System.Windows.Forms.Panel stage_1;
        private System.Windows.Forms.ComboBox cmbPosts;
        private System.Windows.Forms.Label lblPosts;
        private System.Windows.Forms.Label lblGroups;
        private System.Windows.Forms.Label lblNames;
        private System.Windows.Forms.ComboBox cmbGroups;
        private System.Windows.Forms.ComboBox cmbNames;
        private System.Windows.Forms.Label lblQuestionText;
        private System.Windows.Forms.Label lblAnswerD;
        private System.Windows.Forms.Label lblAnswerC;
        private System.Windows.Forms.Label lblAnswerB;
        private System.Windows.Forms.Label lblAnswerA;
        private System.Windows.Forms.Label lblQuestion;
        private System.Windows.Forms.Button btnPrev;
        private System.Windows.Forms.ProgressBar pBar;
        private System.Windows.Forms.Button btnEnd;
        private System.Windows.Forms.Panel stage_2;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.Timer Time;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox YesNoLabel;
    }
}

