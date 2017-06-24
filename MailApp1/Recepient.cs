/*
 * Created by SharpDevelop.
 * User: a0714786
 * Date: 26/01/2017
 * Time: 15:42
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections;

namespace MailApp1
{

	/// <summary>
	/// Description of Recepient.
	/// </summary>
	public class Recepient{
        private String emailAddress;
		private SortedList fieldsToReplace;
        private String cc;
        private String bcc;
        private String subject;
		public SortedList FieldsToReplace {
			get { return fieldsToReplace; }
			set { fieldsToReplace = value; }
		}
		private String employeeID;
		public String EmployeeID {
			get { return employeeID; }
			set { employeeID = value; }
		}

        public string EmailAddress
        {
            get
            {
                return emailAddress;
            }

            set
            {
                emailAddress = value;
            }
        }

        public string Cc
        {
            get
            {
                return cc;
            }

            set
            {
                cc = value;
            }
        }

        public string Bcc
        {
            get
            {
                return bcc;
            }

            set
            {
                bcc = value;
            }
        }

        public string Subject
        {
            get
            {
                return subject;
            }

            set
            {
                subject = value;
            }
        }

        public void MapSubject(String key)
        {
            if(fieldsToReplace.ContainsKey(key)) subject = (String)FieldsToReplace[key];
        }
        public void MapCc(String key)
        {
            if (fieldsToReplace.ContainsKey(key)) Cc = (String)FieldsToReplace[key];
        }

        public void MapBcc(String key)
        {
            if (fieldsToReplace.ContainsKey(key)) Bcc = (String)FieldsToReplace[key];
        }

        public void mapEmailAdress(String key)
        {
            if (fieldsToReplace.ContainsKey(key)) emailAddress = (String)FieldsToReplace[key];
        }
        public void mapEmployeeID(String key)
        {
            if (fieldsToReplace.ContainsKey(key)) employeeID = (String)FieldsToReplace[key];
        }



        public Recepient(SortedList fieldsToMap)
		{
			fieldsToReplace=fieldsToMap;
		}
		
	}
}