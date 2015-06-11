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


        string blueberryAPIurl = "http://localhost.:8080/";
        //public static string blueberryAPIurl = "http://blueberry-api.appspot.com/";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            
        }

        /* Buttons Events */

        private void Publish_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.MyTaskPane.Visible = true;

        }

        private void Fetch_Click(object sender, RibbonControlEventArgs e)
        {
            saveToExcel(fetchData());
        }

        private void Refresh_Click(object sender, RibbonControlEventArgs e)
        {
            Dictionary<string, dynamic> fetchedData = getFetched();
            int fetchDataItemsCount = fetchedData["names"].Count;
            for (int i = 0; i < fetchDataItemsCount; i++)
            {
                Dictionary<string, dynamic> singleResult = new Dictionary<string, dynamic>();
                singleResult.Add("bapi_id", fetchedData["names"][i]);
                singleResult.Add("user", fetchedData["users"][i]);
                singleResult.Add("description", fetchedData["descriptions"][i]);
                singleResult.Add("organization", fetchedData["organizations"][i]);
                singleResult.Add("workbook_path", fetchedData["workbook_paths"][i]);
                singleResult.Add("workbook", fetchedData["workbooks"][i]);
                singleResult.Add("worksheet", fetchedData["worksheets"][i]);
                singleResult.Add("destination_cell", fetchedData["destination_cells"][i]);
                saveToExcel(fetchData(singleResult));
            }
        }

        private void Update_Click(object sender, RibbonControlEventArgs e)
        {
            Dictionary<string, dynamic> publishedData = getPublished();
            int publishedDataItemsCount = publishedData["ids"].Count;

            for (int i = 0; i < publishedDataItemsCount; i++)
            {
                Dictionary<string, dynamic> singleResult = new Dictionary<string, dynamic>();
                singleResult.Add("bapi_id", publishedData["ids"][i]);
                singleResult.Add("user", publishedData["users"][i]);
                singleResult.Add("name", publishedData["names"][i]);
                singleResult.Add("description", publishedData["descriptions"][i]);
                singleResult.Add("organization", publishedData["organizations"][i]);
                singleResult.Add("workbook_path", publishedData["workbook_paths"][i]);
                singleResult.Add("workbook", publishedData["workbooks"][i]);
                singleResult.Add("worksheet", publishedData["worksheets"][i]);
                singleResult.Add("destination_cell", publishedData["destination_cells"][i]);
                singleResult.Add("data_type", publishedData["data_types"][i]);

                publishData(singleResult);

            }
        }

        /* Communication Methods */

        public static void publishData(Dictionary<string, dynamic> singleResult = null)
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
            String xlID;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            ArrayList publishingList = new ArrayList();

            if (singleResult == null)
            {

                xlRange = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                xlPath = xlWorkBook.Path;
                xlWorkbookName = xlWorkBook.Name;
                xlWorksheetName = xlWorkSheet.Name;
                xlDestinationCell = xlRange.Cells[1, 1].Address;
                xlType = "List";
                xlName = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingNameTextBox.Text;
                xlDescription = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingDescriptionTextBox.Text;
                xlOrganization = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingOrganizationTextBox.Text;
                xlDataOwner = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingDataOwnerTextBox.Text;
                xlID = xlOrganization + "." + xlName.Replace(" ", "_") + "." + xlType;

                int rCnt = 0;
                int cCnt = 0;

                for (rCnt = 1; rCnt <= xlRange.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
                    {
                        publishingList.Add((string)(xlRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    }
                }
            }
            else
            {

                Excel.Range xlStartRange;
                Excel.Range xlEndRange;
                xlPath = singleResult["workbook_path"];
                xlWorkbookName = singleResult["workbook"];
                xlWorksheetName = singleResult["worksheet"];
                xlDestinationCell = singleResult["destination_cell"];
                xlType = singleResult["data_type"];
                xlName = singleResult["name"];
                xlDescription = singleResult["description"];
                xlOrganization = singleResult["organization"];
                xlDataOwner = singleResult["user"];
                xlStartRange = (Excel.Range)xlWorkSheet.Range[xlDestinationCell];
                xlEndRange = xlStartRange.End[Excel.XlDirection.xlDown];
                xlRange = (Excel.Range)xlWorkSheet.Range[xlStartRange, xlEndRange];
                xlID = singleResult["bapi_id"];

                int rCnt = 0;
                int cCnt = 0;

                for (rCnt = 1; rCnt <= xlRange.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
                    {
                        publishingList.Add((string)(xlRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    }
                }
            }

            Dictionary<string, dynamic> publishingData = new Dictionary<string, dynamic>();
            publishingData.Add("data", publishingList);
            publishingData.Add("workbook_path", xlPath);
            publishingData.Add("workbook", xlWorkbookName);
            publishingData.Add("worksheet", xlWorksheetName);
            publishingData.Add("destination_cell", xlDestinationCell);
            publishingData.Add("data_type", xlType);
            publishingData.Add("name", xlName);
            publishingData.Add("description", xlDescription);
            publishingData.Add("organization", xlOrganization);
            publishingData.Add("user", xlDataOwner);
            publishingData.Add("bapi_id", xlID);

            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(publishingData);


            var httpWebRequest = (HttpWebRequest)WebRequest.Create(blueberryAPIurl + "List.publish");
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


        }

        private string fetchData(Dictionary<string, dynamic> singleResult = null)
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
            Dictionary<string, dynamic> fetchingData = new Dictionary<string, dynamic>();

            if (singleResult == null)
            {
                xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
                xlRange = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
                xlWorkbookPath = xlWorkBook.Path;
                xlWorkbookName = xlWorkBook.Name;
                xlWorksheetName = xlWorkSheet.Name;
                xlDestinationCell = xlRange.Address;
                xlBlueberryID = IDBox.Text;
                xlDataOwner = "bartosz.piechnik@ch.abb.com";
                xlFetchConfiguration = FetchConfigurationCheckBox.Checked;

                fetchingData.Add("bapi_id", xlBlueberryID);
                fetchingData.Add("user", xlDataOwner);
                fetchingData.Add("workbook_path", xlWorkbookPath);
                fetchingData.Add("workbook", xlWorkbookName);
                fetchingData.Add("worksheet", xlWorksheetName);
                fetchingData.Add("destination_cell", xlDestinationCell);
                fetchingData.Add("skip_new_conf", !xlFetchConfiguration);
            }

            else
            {
                fetchingData.Add("bapi_id", singleResult["bapi_id"]);
                fetchingData.Add("user", singleResult["user"]);
                fetchingData.Add("workbook_path", singleResult["workbook_path"]);
                fetchingData.Add("workbook", singleResult["workbook"]);
                fetchingData.Add("worksheet", singleResult["worksheet"]);
                fetchingData.Add("destination_cell", singleResult["destination_cell"]);
                fetchingData.Add("skip_new_conf", true);
            }

            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(fetchingData);

            String[] splitWords = fetchingData["bapi_id"].Split('.');
            String url = blueberryAPIurl + splitWords[2] + ".fetch";

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
                string result;
                result = streamReader.ReadToEnd();
                return result;

            }
        }

        private void saveToExcel(string result)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;
            String xlWorksheetName;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            var jsonSerializer = new JavaScriptSerializer();

            Dictionary<string, dynamic> fetchedData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(result);

            if (fetchedData["destination_cell"] != null)
            {
                xlRange = (Excel.Range)xlWorkSheet.Range[fetchedData["destination_cell"]];
            }
            else
            {
                xlRange = (Excel.Range)xlWorkSheet.Application.Selection;
            }

            Int32 dataLength = fetchedData["data"].Count;
            Excel.Range endCell = (Excel.Range)xlWorkSheet.Cells[xlRange.Row + dataLength - 1, xlRange.Column];
            Excel.Range xlDestinationRange = xlWorkSheet.Range[xlRange, endCell];

            var fetchedDataArray = new object[dataLength, 1];
            for (var i = 0; i < dataLength; i++)
            {
                fetchedDataArray[i, 0] = fetchedData["data"][i];
            }

            xlDestinationRange.Value2 = fetchedDataArray;
        }


        private Dictionary<string, dynamic> getFetched()
        {
            String xlWorkbookPath;
            String xlWorkbookName;
            Excel.Workbook xlWorkBook;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkbookName = xlWorkBook.Name;
            xlWorkbookPath = xlWorkBook.Path;

            Dictionary<string, dynamic> activeWorkbookInfo = new Dictionary<string, dynamic>();
            activeWorkbookInfo.Add("workbook_path", xlWorkbookPath);
            activeWorkbookInfo.Add("workbook", xlWorkbookName);


            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(activeWorkbookInfo);

            String url = blueberryAPIurl + "List.get_fetched";

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
                string result;
                result = streamReader.ReadToEnd();
                Dictionary<string, dynamic> fetchedData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(result);
                return fetchedData;

            }

        }

        private Dictionary<string, dynamic> getPublished()
        {
            String xlWorkbookPath;
            String xlWorkbookName;
            Excel.Workbook xlWorkBook;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkbookName = xlWorkBook.Name;
            xlWorkbookPath = xlWorkBook.Path;

            Dictionary<string, dynamic> activeWorkbookInfo = new Dictionary<string, dynamic>();
            activeWorkbookInfo.Add("workbook_path", xlWorkbookPath);
            activeWorkbookInfo.Add("workbook", xlWorkbookName);


            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(activeWorkbookInfo);

            String url = blueberryAPIurl + "List.get_published";

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
                string result;
                result = streamReader.ReadToEnd();
                Dictionary<string, dynamic> publishedData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(result);
                return publishedData;
            }
        }

        /* Garbage collections methods */

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

    }
}
