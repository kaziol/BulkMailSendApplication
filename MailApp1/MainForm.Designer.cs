/*
 * Created by SharpDevelop.
 * User: a0714786
 * Date: 24/01/2017
 * Time: 15:55
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace MailApp1
{
	partial class MainForm
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button2 = new System.Windows.Forms.Button();
            this.openFileDialogMsg = new System.Windows.Forms.OpenFileDialog();
            this.label2 = new System.Windows.Forms.Label();
            this.button3 = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.button4 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.button5 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.folderBrowserDialogAttachments = new System.Windows.Forms.FolderBrowserDialog();
            this.listBoxAttachments = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBoxIDs = new System.Windows.Forms.ComboBox();
            this.comboBoxMails = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.comboBoxCC = new System.Windows.Forms.ComboBox();
            this.comboBoxBCC = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.comboBoxOutlookAcc = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.comboBoxSubject = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.button8 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(36, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(280, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Template email directory";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(322, 31);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(206, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Open email template";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Excel files|*.xls*";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(322, 72);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(206, 23);
            this.button2.TabIndex = 3;
            this.button2.Text = "Open Excel File";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2Click);
            // 
            // openFileDialogMsg
            // 
            this.openFileDialogMsg.FileName = "openFileDialogMsg";
            this.openFileDialogMsg.Filter = "Outlook email file|*.msg";
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(36, 51);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(280, 23);
            this.label2.TabIndex = 4;
            this.label2.Text = "Excel directory";
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(322, 109);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(206, 23);
            this.button3.TabIndex = 5;
            this.button3.Text = "Choose output directory";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(36, 93);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Output directory";
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(200, 183);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(258, 66);
            this.button4.TabIndex = 7;
            this.button4.Text = "Generate Messages";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(51, 310);
            this.listBox1.Name = "listBox1";
            this.listBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBox1.Size = new System.Drawing.Size(226, 251);
            this.listBox1.TabIndex = 8;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // webBrowser1
            // 
            this.webBrowser1.Location = new System.Drawing.Point(343, 342);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(278, 218);
            this.webBrowser1.TabIndex = 9;
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(322, 148);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(206, 23);
            this.button5.TabIndex = 10;
            this.button5.Text = "Choose attachment directory";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(36, 134);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(104, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Attachment directory";
            // 
            // listBoxAttachments
            // 
            this.listBoxAttachments.FormattingEnabled = true;
            this.listBoxAttachments.Location = new System.Drawing.Point(643, 310);
            this.listBoxAttachments.Name = "listBoxAttachments";
            this.listBoxAttachments.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.listBoxAttachments.Size = new System.Drawing.Size(226, 251);
            this.listBoxAttachments.TabIndex = 12;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(340, 282);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(43, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Subject";
            // 
            // textBox1
            // 
            this.textBox1.Enabled = false;
            this.textBox1.Location = new System.Drawing.Point(39, 28);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(265, 20);
            this.textBox1.TabIndex = 14;
            // 
            // textBox2
            // 
            this.textBox2.Enabled = false;
            this.textBox2.Location = new System.Drawing.Point(39, 70);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(265, 20);
            this.textBox2.TabIndex = 15;
            // 
            // textBox3
            // 
            this.textBox3.Enabled = false;
            this.textBox3.Location = new System.Drawing.Point(39, 109);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(265, 20);
            this.textBox3.TabIndex = 16;
            // 
            // textBox4
            // 
            this.textBox4.Enabled = false;
            this.textBox4.Location = new System.Drawing.Point(39, 151);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(265, 20);
            this.textBox4.TabIndex = 17;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(343, 298);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(265, 20);
            this.textBox5.TabIndex = 18;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(640, 282);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 13);
            this.label6.TabIndex = 19;
            this.label6.Text = "Attachment List";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(340, 321);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(31, 13);
            this.label7.TabIndex = 20;
            this.label7.Text = "Body";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(48, 282);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(237, 13);
            this.label8.TabIndex = 21;
            this.label8.Text = "Generated output files to the following recepients";
            // 
            // comboBoxIDs
            // 
            this.comboBoxIDs.FormattingEnabled = true;
            this.comboBoxIDs.Location = new System.Drawing.Point(571, 207);
            this.comboBoxIDs.Name = "comboBoxIDs";
            this.comboBoxIDs.Size = new System.Drawing.Size(121, 21);
            this.comboBoxIDs.TabIndex = 22;
            // 
            // comboBoxMails
            // 
            this.comboBoxMails.FormattingEnabled = true;
            this.comboBoxMails.Location = new System.Drawing.Point(571, 85);
            this.comboBoxMails.Name = "comboBoxMails";
            this.comboBoxMails.Size = new System.Drawing.Size(121, 21);
            this.comboBoxMails.TabIndex = 23;
            // 
            // label9
            // 
            this.label9.Location = new System.Drawing.Point(568, 60);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(205, 35);
            this.label9.TabIndex = 24;
            this.label9.Text = "Select Excel Column as recepients email address";
            // 
            // label10
            // 
            this.label10.Location = new System.Drawing.Point(568, 178);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(188, 40);
            this.label10.TabIndex = 25;
            this.label10.Text = "Select Excel Column as attachment identifier";
            // 
            // label11
            // 
            this.label11.Location = new System.Drawing.Point(568, 114);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(153, 50);
            this.label11.TabIndex = 27;
            this.label11.Text = "Select Excel Column as CC identifier";
            // 
            // comboBoxCC
            // 
            this.comboBoxCC.FormattingEnabled = true;
            this.comboBoxCC.Location = new System.Drawing.Point(571, 143);
            this.comboBoxCC.Name = "comboBoxCC";
            this.comboBoxCC.Size = new System.Drawing.Size(121, 21);
            this.comboBoxCC.TabIndex = 26;
            // 
            // comboBoxBCC
            // 
            this.comboBoxBCC.FormattingEnabled = true;
            this.comboBoxBCC.Location = new System.Drawing.Point(782, 143);
            this.comboBoxBCC.Name = "comboBoxBCC";
            this.comboBoxBCC.Size = new System.Drawing.Size(121, 21);
            this.comboBoxBCC.TabIndex = 28;
            // 
            // label12
            // 
            this.label12.Location = new System.Drawing.Point(779, 114);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(153, 50);
            this.label12.TabIndex = 29;
            this.label12.Text = "Select Excel Column as BCC identifier";
            // 
            // comboBoxOutlookAcc
            // 
            this.comboBoxOutlookAcc.FormattingEnabled = true;
            this.comboBoxOutlookAcc.Location = new System.Drawing.Point(571, 33);
            this.comboBoxOutlookAcc.Name = "comboBoxOutlookAcc";
            this.comboBoxOutlookAcc.Size = new System.Drawing.Size(332, 21);
            this.comboBoxOutlookAcc.TabIndex = 30;
            // 
            // label13
            // 
            this.label13.Location = new System.Drawing.Point(568, 19);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(205, 35);
            this.label13.TabIndex = 31;
            this.label13.Text = "Select Outlook Account";
            // 
            // comboBoxSubject
            // 
            this.comboBoxSubject.FormattingEnabled = true;
            this.comboBoxSubject.Location = new System.Drawing.Point(782, 85);
            this.comboBoxSubject.Name = "comboBoxSubject";
            this.comboBoxSubject.Size = new System.Drawing.Size(121, 21);
            this.comboBoxSubject.TabIndex = 32;
            // 
            // label14
            // 
            this.label14.Location = new System.Drawing.Point(779, 60);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(205, 35);
            this.label14.TabIndex = 33;
            this.label14.Text = "Select Excel Column as Subject";
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(39, 586);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(238, 48);
            this.button6.TabIndex = 34;
            this.button6.Text = "Send selected mails";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(343, 586);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(234, 48);
            this.button7.TabIndex = 35;
            this.button7.Text = "Send All";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(724, 225);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(179, 45);
            this.progressBar1.TabIndex = 36;
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(857, 195);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(75, 23);
            this.button8.TabIndex = 37;
            this.button8.Text = "button8";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(964, 646);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.comboBoxSubject);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.comboBoxOutlookAcc);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.comboBoxBCC);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.comboBoxIDs);
            this.Controls.Add(this.comboBoxCC);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.comboBoxMails);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.listBoxAttachments);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Name = "MainForm";
            this.Text = "SPAM Application";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		public System.Windows.Forms.Label label2;
		private System.Windows.Forms.OpenFileDialog openFileDialogMsg;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.Button button1;
		public System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogAttachments;
        private System.Windows.Forms.ListBox listBoxAttachments;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBoxIDs;
        private System.Windows.Forms.ComboBox comboBoxMails;
        public System.Windows.Forms.Label label9;
        public System.Windows.Forms.Label label10;
        public System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox comboBoxCC;
        private System.Windows.Forms.ComboBox comboBoxBCC;
        public System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox comboBoxOutlookAcc;
        public System.Windows.Forms.Label label13;
        private System.Windows.Forms.ComboBox comboBoxSubject;
        public System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button button8;
    }
}
