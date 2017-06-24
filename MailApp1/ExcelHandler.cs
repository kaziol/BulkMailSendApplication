/*
 * Created by SharpDevelop.
 * User: a0714786
 * Date: 24/01/2017
 * Time: 17:21
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Collections.Generic;
using System.Windows.Forms;

namespace MailApp1
{
	/// <summary>
	/// Description of ExcelHandler.
	/// </summary>
	public class ExcelHandler
	{
		public String name;
		private List<String> headers;
		private List<Recepient> recepients;
		Excel.Workbook workbook;
        private Boolean errorFree;
		public ExcelHandler(String filePath)
		{
			Excel.Application excelApp = new Excel.Application();
            ErrorFree = true;
            workbook =excelApp.Workbooks.Open(filePath);
			name=workbook.Name;
			Headers=new List<string>();
			Recepients=new List<Recepient>();
			retrieveDataFromExcel();
		}

        public List<Recepient> Recepients
        {
            get
            {
                return recepients;
            }

            set
            {
                recepients = value;
            }
        }

        public bool ErrorFree
        {
            get
            {
                return errorFree;
            }

            set
            {
                errorFree = value;
            }
        }

        public List<string> Headers
        {
            get
            {
                return headers;
            }

            set
            {
                headers = value;
            }
        }

        private void retrieveDataFromExcel(){
			try{
			Excel.Worksheet xlWorksheet = (Excel.Worksheet) workbook.Sheets[1];
			Excel.Range xlRange = xlWorksheet.UsedRange;
			int rowCount=xlRange.Rows.Count;
			int columnCount = xlRange.Columns.Count;
              //  String emailHeader =null, empIDHeader =null;
			for (int i=1;i<rowCount+1;i++){
				SortedList fieldMap = new SortedList();
				for(int j=1;j<columnCount+1;j++){
					object rangeObject = xlWorksheet.Cells[i, j];
					Excel.Range range = (Excel.Range) rangeObject;
					Object rangeValue = range.Value2;

                        String cellValue;
                        try
                        {
                            cellValue = rangeValue.ToString();
                        }catch(NullReferenceException)
                        {
                            cellValue = String.Empty;
                        }		
					if(i==1)
                        {
						Headers.Add(cellValue);
                    }
					else
                        {
						fieldMap.Add(Headers[j-1],cellValue);
					}
				}
				if(i!=1)Recepients.Add(new Recepient(fieldMap));

			}
			}catch(Exception ex){
                MessageBox.Show(ex.Message, "Error while reading Excel file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ErrorFree = false;
			}
			finally{
				workbook.Close();
			}					
		}
	}
}
