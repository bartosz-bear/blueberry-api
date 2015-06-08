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
using System.Data;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {


        public static string blueberryAPIurl = "http://localhost:8080/";
        //public static string blueberryAPIurl = "http://blueberry-api.appspot.com/";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            MessageBox.Show("Bartosz Piechnik");
        }

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            MessageBox.Show("Bartosz");
        }

        /* Buttons Events */

        private void Publish_Click(object sender, RibbonControlEventArgs e)
        {
            //BlueberryTaskPane publishBlueberryTaskPane;
            //Microsoft.Office.Tools.CustomTaskPane myTaskPane;

            //Excel.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(
            //    Application_WorkbookActivate);

            //MessageBox.Show("Clicked 2");

            //publishBlueberryTaskPane = new BlueberryTaskPane();
            //myTaskPane = this.CustomTaskPanes.Add(
            //    publishBlueberryTaskPane, "Publish");

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
                singleResult.Add("bapi_id", fetchedData["bapi_ids"][i]);
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

            Dictionary<string, dynamic> dataFromExcel = new Dictionary<string, dynamic>();
            

            if (singleResult == null)
            {

                xlRange = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                xlPath = xlWorkBook.Path;
                xlWorkbookName = xlWorkBook.Name;
                xlWorksheetName = xlWorkSheet.Name;
                xlDestinationCell = xlRange.Cells[1, 1].Address;
                xlName = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingNameTextBox.Text;
                xlDescription = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingDescriptionTextBox.Text;
                xlOrganization = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingOrganizationTextBox.Text;
                xlDataOwner = Globals.ThisAddIn.publishBlueberryTaskPane.PublishingDataOwnerTextBox.Text;
                dataFromExcel = measureData(xlRange);
                xlType = dataFromExcel["data_type"];
                xlID = xlOrganization + "." + xlName.Replace(" ", "_") + "." + xlType;

                /*
                int rCnt = 0;
                int cCnt = 0;

                for (rCnt = 1; rCnt <= xlRange.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
                    {
                        publishingList.Add((string)(xlRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    }
                }
                */
            }
            else
            {


                xlPath = singleResult["workbook_path"];
                xlWorkbookName = singleResult["workbook"];
                xlWorksheetName = singleResult["worksheet"];
                xlDestinationCell = singleResult["destination_cell"];
                xlType = singleResult["data_type"];
                xlName = singleResult["name"];
                xlDescription = singleResult["description"];
                xlOrganization = singleResult["organization"];
                xlDataOwner = singleResult["user"];
                
                // Here we will need a separate function which will help to establish a Range depending on the data type
                xlRange = specifyRange(xlWorkSheet, xlDestinationCell);
                xlID = singleResult["bapi_id"];

                dataFromExcel = measureData(xlRange);
                dataFromExcel["data_type"] = xlType;

                /*
                int rCnt = 0;
                int cCnt = 0;

                for (rCnt = 1; rCnt <= xlRange.Rows.Count; rCnt++)
                {
                    for (cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
                    {
                        publishingList.Add((string)(xlRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    }
                }
                */
            }

            Dictionary<string, dynamic> publishingData = new Dictionary<string, dynamic>();
            publishingData.Add("data", dataFromExcel["data"]);
            publishingData.Add("workbook_path", xlPath);
            publishingData.Add("workbook", xlWorkbookName);
            publishingData.Add("worksheet", xlWorksheetName);
            publishingData.Add("destination_cell", xlDestinationCell);
            publishingData.Add("data_type", dataFromExcel["data_type"]);
            publishingData.Add("name", xlName);
            publishingData.Add("description", xlDescription);
            publishingData.Add("organization", xlOrganization);
            publishingData.Add("user", xlDataOwner);
            publishingData.Add("bapi_id", xlID);

            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(publishingData);


            var httpWebRequest = (HttpWebRequest)WebRequest.Create(blueberryAPIurl + xlType + ".publish");
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


        
        public static Dictionary<string, dynamic> measureData(Excel.Range xlRange)
        {
            int xlRowsCount;
            int xlColumnsCount;
            xlRowsCount = xlRange.Rows.Count;
            xlColumnsCount = xlRange.Columns.Count;

            Dictionary<string, dynamic> dataInfo = new Dictionary<string, dynamic>();
            dataInfo.Add("rows_count", xlRowsCount);
            dataInfo.Add("columns_count", xlColumnsCount);
            dataInfo.Add("data_type", labelData(xlRowsCount, xlColumnsCount));
            dataInfo.Add("data", fromExcelToObject(xlRowsCount, xlColumnsCount, dataInfo["data_type"], xlRange));
            return dataInfo;
        }

        public static String labelData(int xlRowsCount, int xlColumnsCount)
        {
            switch (xlColumnsCount)
            {
                case 1:
                    if (xlRowsCount == 1)
                    {
                        return "Scalar";
                    }
                    else
                    {
                        return "List";
                    }
                case 2:
                    return "Dictionary";
                default:
                    if (xlColumnsCount > 2) {
                        return "Table";
                    }
                    else
                    {
                        return "Other";
                    }
            }
        }
        

        public static dynamic fromExcelToObject(int rowsCount, int columnsCount, string dataType, Excel.Range xlRange)
        {
            switch (dataType)
            {
                case "Scalar":
                    return xlRange.Value2;
                case "List":
                    ArrayList publishingList = new ArrayList();
                    for (int currentRowsCount = 1; currentRowsCount <= rowsCount; currentRowsCount++)
                    {
                        for (int currentColumnsCount = 1; currentColumnsCount <= columnsCount; currentColumnsCount++)
                        {
                            publishingList.Add((dynamic)(xlRange.Cells[currentRowsCount, currentColumnsCount] as Excel.Range).Value2);
                        }
                    }
                    return publishingList;
                case "Dictionary":
                    Dictionary<string, dynamic> publishingDictionary = new Dictionary<string,dynamic>();
                    for (int currentRowsCount = 1; currentRowsCount <= rowsCount; currentRowsCount++)
                    {
                        publishingDictionary.Add((dynamic)(xlRange.Cells[currentRowsCount, 1] as Excel.Range).Value2,
                                                 (dynamic)(xlRange.Cells[currentRowsCount, 2] as Excel.Range).Value2);
                    }
                    return publishingDictionary;
                case "Table":
                    List<List<object>> publishingTable = new List<List<object>>();
                    int columnsCountCopy = columnsCount;
                    for (int currentColumnsCount = 1; currentColumnsCount <= columnsCount; currentColumnsCount++)
                    {
                        //Type sublistType = (dynamic)(xlRange.Cells[1, currentColumnsCount] as Excel.Range).Value2.GetType();
                        List<object> sublist = new List<object>();
                        for (int currentRowsCount = 1; currentRowsCount <= rowsCount; currentRowsCount++)
                        {
                            sublist.Add((dynamic)(xlRange.Cells[currentRowsCount, currentColumnsCount] as Excel.Range).Value2);
                        }
                        publishingTable.Add(sublist);
                    }

                    return publishingTable;
                default:
                    return "Other";
            }
            

        }

        public static Excel.Range specifyRange(Excel.Worksheet xlWorkSheet, string xlDestinationCell)
        {
            Excel.Range xlStartRange;
            Excel.Range xlEndRange;
            Excel.Range xlRange;
            xlStartRange = (Excel.Range)xlWorkSheet.Range[xlDestinationCell];

            switch (xlDestinationCell)
            {
                case "Scalar":
                    xlEndRange = xlStartRange;
                    break;
                case "List":
                    xlEndRange = xlStartRange.End[Excel.XlDirection.xlDown];
                    break;
                case "Dictionary":
                    xlEndRange = xlStartRange.End[Excel.XlDirection.xlDown].End[Excel.XlDirection.xlToRight];
                    break;
                case "Table":
                    xlEndRange = xlStartRange.End[Excel.XlDirection.xlDown].End[Excel.XlDirection.xlToRight];
                    break;
                default:
                    xlEndRange = xlStartRange;
                    break;
            }

            xlRange = (Excel.Range)xlWorkSheet.Range[xlStartRange, xlEndRange];
            return xlRange;
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

            // This List.get_fetched should be abstracted to accomodate Scalar, List, Dict, etc.
            // This function should check all different possible data structures, because a single workbook can have
            // different data structures.
            String url = blueberryAPIurl + "Data.get_fetched";

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

        public static Boolean isPublishRangeEmpty()
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            xlRange = (Excel.Range)xlWorkSheet.Application.Selection;

            if (xlRange.Value2 == null)
            {
                return true;
            }
            else
            {
                return false;
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

            // This List.get_published should be abstracted to accomodate Scalar, List, Dict, etc.
            // This function should check all different possible data structures, because a single workbook can have
            // different data structures.
            String url = blueberryAPIurl + "Data.get_published";

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

        private void TestButton_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Testing");
        }

    }
}
