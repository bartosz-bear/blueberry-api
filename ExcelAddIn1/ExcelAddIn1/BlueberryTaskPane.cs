using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Web.Script.Serialization;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace ExcelAddIn1
{
    public partial class BlueberryTaskPane : UserControl
    {
        public BlueberryTaskPane()
        {
            InitializeComponent();
        }

        private void PublishButton_Click(object sender, EventArgs e)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;
            String xlPath;
            String xlWorkbookName;
            String xlWorksheetName;
            String xlDestinationCell;
            String xlType;
            String xlName;
            String xlDescription;
            String xlOrganization;
            String xlDataOwner;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            xlRange = xlWorkSheet.UsedRange;
            xlPath = xlWorkBook.Path;
            xlWorkbookName = xlWorkBook.Name;
            xlWorksheetName = xlWorkSheet.Name;
            xlDestinationCell = xlRange.Cells[1, 1].Address;
            xlType = "List";
            xlName = NameTextBox.Text;
            xlDescription = DescriptionTextBox.Text;
            xlOrganization = OrganizationTextBox.Text;
            xlDataOwner = DataOwnerTextBox.Text;

            ArrayList publishingList = new ArrayList();

            int rCnt = 0;
            int cCnt = 0;

            for (rCnt = 1; rCnt <= xlRange.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
                {
                    publishingList.Add((string)(xlRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                }
            }

            Dictionary<string, dynamic> publishingData = new Dictionary<string, dynamic>();
            publishingData.Add("aq_data", publishingList);
            publishingData.Add("aq_workbook_path", xlPath);
            publishingData.Add("aq_workbook", xlWorkbookName);
            publishingData.Add("aq_worksheet", xlWorksheetName);
            publishingData.Add("aq_destination_cell", xlDestinationCell);
            publishingData.Add("aq_type", xlType);
            publishingData.Add("aq_name", xlName);
            publishingData.Add("aq_description", xlDescription);
            publishingData.Add("aq_organization", xlOrganization);
            publishingData.Add("aq_created_by", xlDataOwner);

            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(publishingData);

            var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://localhost:8080/List.publish");
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();
            }

            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
                MessageBox.Show(result.ToString());
            }

            string str;
            str = xlRange.Value2;
            MessageBox.Show(str);
            string str2;
            str2 = xlRange.Value2;
            MessageBox.Show(str2);
        }
    }
}
