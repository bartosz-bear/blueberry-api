using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Net;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //String debugging = "";
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;

            Globals.ThisAddIn.MyTaskPane.Visible = true;

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            xlRange = xlWorkSheet.UsedRange;



            ArrayList publishingList = new ArrayList();

            //string str;
            int rCnt = 0;
            int cCnt = 0;

            for (rCnt = 1; rCnt <= xlRange.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
                {
                    publishingList.Add((string)(xlRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    //str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    //MessageBox.Show(str);
                }
            }

            MessageBox.Show((string)publishingList[0]);
            MessageBox.Show((string)publishingList[1]);
            MessageBox.Show(rCnt.ToString());

            /*

            string str;
            str = xlRange.Value2;
            MessageBox.Show(str);
            string str2;
            str2 = xlRange.Value2;
            MessageBox.Show(str2);
            
            

            */




            /// This is a working code which creates a new file and saves data
            /*
            excel.application xlapp = new microsoft.office.interop.excel.application();

            if (xlapp == null)
            {
                messagebox.show("excel is not properly installed!!");
                return;
            }

            
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Sheet 1 content";

            xlWorkBook.SaveAs("C:\\Users\\chbapie\\Desktop\\Bartosz\\apiquitous\\spreadsheets\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xls");
            */

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.MyTaskPane.Visible = true;

        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {

            //Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;
            //Microsoft.Office.Tools.CustomTaskPane tempPane = Globals.ThisAddIn.MyTaskPane;
            
            //Globals.ThisAddIn.MyTaskPane.Visible = true;

            MessageBox.Show("Clicked");
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;
            String xlWorkbookPath;
            String xlWorkbookName;
            String xlWorksheetName;
            String xlDestinationCell;
            String xlBlueberryID;
            String xlDataOwner;
            Boolean xlFetchConfiguration;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            xlRange = xlWorkSheet.UsedRange;
            xlWorkbookPath = xlWorkBook.Path;
            xlWorkbookName = xlWorkBook.Name;
            xlWorksheetName = xlWorkSheet.Name;
            xlDestinationCell = xlRange.Address;
            xlBlueberryID = IDBox.Text;
            xlDataOwner = "bartosz.piechnik@ch.abb.com";
            xlFetchConfiguration = FetchConfigurationCheckBox.Checked;

            Dictionary<string, dynamic> fetchingData = new Dictionary<string, dynamic>();
            fetchingData.Add("aq_id", xlBlueberryID);
            fetchingData.Add("user", xlDataOwner);
            fetchingData.Add("workbook_path", xlWorkbookPath);
            fetchingData.Add("workbook", xlWorkbookName);
            fetchingData.Add("worksheet", xlWorksheetName);
            fetchingData.Add("destination_cell", xlDestinationCell);
            fetchingData.Add("skip_new_conf", xlFetchConfiguration);

            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(fetchingData);

            String[] splitWords = xlBlueberryID.Split('.');
            String url = "http://localhost:8080/" + splitWords[2] + ".fetch";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
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
                xlRange = (Excel.Range)xlWorkSheet.Application.Selection;
                Dictionary<string, dynamic> fetchedData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(result);
                Int32 dataLength = fetchedData["aq_data"].Count;
                Excel.Range endCell = (Excel.Range)xlWorkSheet.Cells[xlRange.Row + dataLength - 1, xlRange.Column];
                Excel.Range xlDestinationRange = xlWorkSheet.Range[xlRange, endCell];

                var fetchedDataArray = new object[dataLength, 1];
                for (var i = 0; i < dataLength; i++)
                {
                    fetchedDataArray[i, 0] = fetchedData["aq_data"][i];
                }

                xlDestinationRange.Value2 = fetchedDataArray;
            }

        } 
    }
}
