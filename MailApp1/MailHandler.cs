/*
 * Created by SharpDevelop.
 * User: a0714786
 * Date: 24/01/2017
 * Time: 15:58
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailApp1
{
	public class MailHandler
	{
        private Outlook.MailItem mailItem;
        private List<Outlook.MailItem> mailList;
        private Outlook.Application outlookApp;

        public MailHandler(String filePath)
		{
            if (outlookApp == null)
            {
                outlookApp = new Outlook.Application();
            }
            try
            {
                mailItem = (Outlook.MailItem)outlookApp.Session.OpenSharedItem(filePath);
                MailList = new List<Microsoft.Office.Interop.Outlook.MailItem>();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Error while reading msg sample file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
      	
		}

        public List<Outlook.MailItem> MailList
        {
            get
            {
                return mailList;
            }

            set
            {
                mailList = value;
            }
        }

        public int createOutputFile(Recepient recepient, String outputDir, String attachmentDirectory)
        {
            int createdFiles = 0;
            try
            {
                Outlook.MailItem result = (Outlook.MailItem)mailItem.Copy();
                String htmlBody = result.HTMLBody;
                String pattern = "<span class=SpellE>.*</span>";
                String toReplace;
                foreach (Match m in Regex.Matches(htmlBody, pattern))
                {
                    toReplace = m.Value.Replace("<span class=SpellE>",String.Empty).Replace("</span>", String.Empty);
                    htmlBody=htmlBody.Replace(m.Value, toReplace);
                }
                String subject = recepient.Subject != null ? recepient.Subject : result.Subject;
                foreach (string key in recepient.FieldsToReplace.Keys)
                {
                    string mappedValue =(string) recepient.FieldsToReplace[key];
                    htmlBody =htmlBody.Replace(key, mappedValue);
                   // result.Subject=result.Subject.Replace(key, (string)recepient.FieldsToReplace[key]);
                    subject=subject.Replace(key, mappedValue);
                }
                result.HTMLBody = htmlBody;
                result.Subject = subject;
                String resultPath = outputDir + "\\" + recepient.EmployeeID + ".msg";
                result.To = recepient.EmailAddress;
                if (recepient.Cc != null)
                {
                    result.CC = recepient.Cc;
                }
                if (recepient.Bcc != null)
                {
                    result.BCC = recepient.Bcc;
                }
                if (attachmentDirectory != null)
                {
                    string[] fileArray = Directory.GetFiles(attachmentDirectory);
                    foreach (String filename in fileArray)
                    {
                        if (filename.Contains(recepient.EmployeeID))
                        {
                            result.Attachments.Add(filename);
                        }
                    }
                }
               // if(acc!=null) result.SendUsingAccount = acc;
                //result.Sender = acc.DisplayName;
                MailList.Add(result);
                result.SaveAs(resultPath);
                createdFiles++;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error creating Excel file for " +recepient.EmployeeID, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return createdFiles;
            }
            return createdFiles;
        }
        public void sendAll(Outlook.Account acc)
        {
            Hashtable errorsTable = new Hashtable();
            String errorMessage = String.Empty;
            foreach (Outlook.MailItem mailItem in MailList)
            {
                try
                {
                    mailItem.SendUsingAccount = acc;
                    //mailItem.SenderEmailAddress=acc.
                    mailItem.Send();
                }
                catch (Exception ex)
                {
                    if (!errorsTable.ContainsKey(mailItem.To))
                    {
                        errorsTable.Add(mailItem.To, ex.Message);
                    }
                }       
            }
            if (errorsTable.Count > 0)
            {
                foreach (string key in errorsTable.Keys)
                {
                    errorMessage += key + ":\t" + errorsTable[key] + "\n";
                }
            }
            if (errorMessage.Length > 0)
            {
                MessageBox.Show(errorMessage, "Error occured while sending messages", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Message sent", "Messages sent", MessageBoxButtons.OK);
            }

        }
        public void sendSelectedMails(ListBox.SelectedObjectCollection mailAdresses, Outlook.Account acc)
        {
            Hashtable errorsTable = new Hashtable();
            String errorMessage = String.Empty;
            foreach (Outlook.MailItem mail in mailList)
            {
                try
                {
                    if (mailAdresses.Contains((String)mail.To))
                    {
                        mail.SendUsingAccount = acc;
                        mail.Send();
                    }
                }catch(Exception ex)
                {
                    if (mailItem!=null && !errorsTable.ContainsKey(mailItem.To))
                    {
                        errorsTable.Add(mailItem.To, ex.Message);
                    }
                }
                
            }
            if (errorsTable.Count > 0)
            {
                foreach (string key in errorsTable.Keys)
                {
                    errorMessage += key + ":\t" + errorsTable[key] + "\n";
                }
            }
            if (errorMessage.Length > 0)
            {
                MessageBox.Show(errorMessage, "Error occured while sending messages", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Messages sent", "Messages sent", MessageBoxButtons.OK);
            }
            
        }
	}
}
