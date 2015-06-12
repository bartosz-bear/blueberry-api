using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Net;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using BlueberryTaskPane = ExcelAddIn1.BlueberryTaskPane;
using BlueberryRibbon = ExcelAddIn1.Ribbon1;

namespace ExcelAddIn1.Controllers
{
    class Fetching
    {
        public static string fetchData(Dictionary<string, dynamic> singleResult = null)
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
                xlBlueberryID = Globals.Ribbons.Ribbon1.IDBox.Text;
                xlDataOwner = "bartosz.piechnik@ch.abb.com";
                xlFetchConfiguration = Globals.Ribbons.Ribbon1.FetchConfigurationCheckBox.Checked;

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
            String url = BlueberryRibbon.blueberryAPIurl + splitWords[2] + ".fetch";

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



        public static Dictionary<string, dynamic> getFetched()
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
            String url = BlueberryRibbon.blueberryAPIurl + "Data.get_fetched";

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
    }
}
