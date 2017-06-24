/*
 * Created by SharpDevelop.
 * User: a0714786
 * Date: 24/01/2017
 * Time: 15:55
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailApp1
{
    /// <summary>
    /// Description of MainForm.
    /// </summary>
    public partial class MainForm : Form
    {
        MailHandler mailHandler;
        List<Recepient> recepients;
        Outlook.Accounts accounts;
        String outputDirectory;
        String attachmentDirectory;
        const String asTemplate = "Leave as in template";
        Boolean templateSelected , folderSelected , excelSelected;
        Hashtable outlookAccs;
        Outlook.Application outlookApp;
        public MainForm()
        {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            //progressBar1.Maximum = 100;
            templateSelected = false; folderSelected = false; excelSelected = false;
            button4.Enabled = false;
            comboBoxCC.Items.Add(asTemplate);
            comboBoxBCC.Items.Add(asTemplate);
            comboBoxSubject.Items.Add(asTemplate);
            outlookApp = new Outlook.Application();
            accounts = outlookApp.Session.Accounts;
            outlookAccs = new Hashtable();
            foreach(Outlook.Account acc in accounts)
            {
                outlookAccs.Add(acc.DisplayName, acc);
                comboBoxOutlookAcc.Items.Add(acc.DisplayName);
            }
            if (comboBoxOutlookAcc.Items.Count > 0)
            {
                comboBoxOutlookAcc.SelectedIndex = 0;
            }
            
        }
        void turnOnGenerateButton()
        {
            if (templateSelected && folderSelected && excelSelected) button4.Enabled = true;
        }
        void Button1Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (openFileDialogMsg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                String msgFilePath = openFileDialogMsg.FileName;
                mailHandler = new MailHandler(msgFilePath);
                textBox1.Text = msgFilePath;
                templateSelected = true;
                turnOnGenerateButton();
            }
            Cursor.Current = Cursors.Default;
        }



        void Button2Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                String excelName = openFileDialog1.FileName;
                ExcelHandler ex = new ExcelHandler(excelName);
                if (ex.ErrorFree)
                {
                    recepients = ex.Recepients;
                    foreach(String header in ex.Headers)
                    {
                        comboBoxIDs.Items.Add(header);
                        comboBoxMails.Items.Add(header);
                        comboBoxCC.Items.Add(header);
                        comboBoxBCC.Items.Add(header);
                        comboBoxSubject.Items.Add(header);
                    }
                    comboBoxMails.SelectedIndex = 0;
                    comboBoxIDs.SelectedIndex = 1;
                    comboBoxBCC.SelectedIndex = 0;
                    comboBoxCC.SelectedIndex = 0;
                    comboBoxSubject.SelectedIndex = 0;
                    textBox2.Text = excelName;
                    excelSelected = true;
                    turnOnGenerateButton();
                }
            }
            Cursor.Current = Cursors.Default;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (Microsoft.Office.Interop.Outlook.MailItem mailItem in mailHandler.MailList)
            {
                if (mailItem.To == (string) listBox1.SelectedItem)
                {
                    webBrowser1.DocumentText = mailItem.HTMLBody;
                    listBoxAttachments.Items.Clear();
                    foreach (Microsoft.Office.Interop.Outlook.Attachment att in mailItem.Attachments)
                    {
                        listBoxAttachments.Items.Add(att.FileName);
                    }
                    textBox5.Text = mailItem.Subject;
                    break;

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                outputDirectory = folderBrowserDialog1.SelectedPath;
                textBox3.Text = outputDirectory;
                folderSelected = true;
                turnOnGenerateButton();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialogAttachments.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                attachmentDirectory = folderBrowserDialogAttachments.SelectedPath;
                textBox4.Text = attachmentDirectory;
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //selected send
            String outlookAccString = (String)comboBoxOutlookAcc.SelectedItem;
            Outlook.Account acc = (Outlook.Account)outlookAccs[outlookAccString];
            mailHandler.sendSelectedMails(listBox1.SelectedItems, acc);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            String outlookAccString = (String)comboBoxOutlookAcc.SelectedItem;
            Outlook.Account acc = (Outlook.Account)outlookAccs[outlookAccString];
            mailHandler.sendAll(acc);
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            int filesCreatedCount = 0;
            List<object> args = e.Argument as List<object>;
            List<Recepient> recepients = (List<Recepient>)args[0];
            String selectedID = (String)args[1];
            String selectedMail = (String)args[2];
            String selectedCC = (String)args[3];
            String selectedBCC = (String)args[4];
            String selectedSubject = (String)args[5];
            String asTemplate = (String)args[6];
            // backgroundWorker1.RunWorkerAsync(recepients);
            foreach (Recepient recepient in recepients)
            {
                // button4.Enabled = false;
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                recepient.mapEmailAdress(selectedMail);
                recepient.mapEmployeeID(selectedID);
                //null means as in template
                if (selectedCC != asTemplate)
                {
                    recepient.MapCc(selectedCC);
                }
                if (selectedBCC != asTemplate)
                {
                    recepient.MapBcc(selectedBCC);
                }
                if (selectedSubject != asTemplate)
                {
                    recepient.MapSubject(selectedSubject);
                }
                filesCreatedCount += mailHandler.createOutputFile(recepient, outputDirectory, attachmentDirectory);
                backgroundWorker1.ReportProgress(filesCreatedCount * 100 / recepients.Count);
            }   
            e.Result = filesCreatedCount;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                button4.Enabled = true;
                MessageBox.Show("Succesfully created " + e.Result + " files.", "Operation Completed", MessageBoxButtons.OK);
                updateListBox();
            }

        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            progressBar1.Value=e.ProgressPercentage;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            // clear mailList
            mailHandler.MailList= new List<Microsoft.Office.Interop.Outlook.MailItem>();
            List<object> argumentList = new List<object>();
            String selectedID = (String)comboBoxIDs.SelectedItem;
            String selectedMail = (String)comboBoxMails.SelectedItem;
            String selectedCC = (String)comboBoxCC.SelectedItem;
            String selectedBCC = (String)comboBoxBCC.SelectedItem;
            String selectedSubject = (String)comboBoxSubject.SelectedItem;
            argumentList.Add(recepients);
            argumentList.Add(selectedID);
            argumentList.Add(selectedMail);
            argumentList.Add(selectedCC);
            argumentList.Add(selectedBCC);
            argumentList.Add(selectedSubject);
            argumentList.Add(asTemplate);

            backgroundWorker1.RunWorkerAsync(argumentList);
           // updateListBox();
            Cursor.Current = Cursors.Default;
           
        }

        public void updateListBox()
        {
            listBox1.Items.Clear();
            foreach(Microsoft.Office.Interop.Outlook.MailItem mailItem in mailHandler.MailList)
            {
                listBox1.Items.Add(mailItem.To);
            }
        }
    }
}
