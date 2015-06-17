using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Net;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using BlueberryTaskPane = ExcelAddIn1.BlueberryTaskPane;
using BlueberryRibbon = ExcelAddIn1.BlueberryRibbon;

namespace ExcelAddIn1.Controllers
{
    /// <summary>
    /// Fetching class is used to make requests to Blubeberry cloud and fetch BAPI data structures
    /// stored in the Blueberry datastore.
    /// </summary>
    class Fetching
    {
        /// <summary>
        /// fetchData() method makes a request to Blueberry cloud and retrieves a serialized data.
        /// Depending whether data is fetched for the first time or it has been previously fetched
        /// the method populates fetchingData dictionary with data from the active Excel spreadsheet
        /// or from FetchedConfigurations Blueberry cloud datastore class.
        /// </summary>
        /// <param name="singleResult">If data is fetched for the first time then this parameter is null,
        /// otherwise it will be a single record from FetchConfigurations class.</param>
        /// <returns></returns>
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

            // Data is fetched for the first time.
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

            // Data has been fetched before.
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

            // Serializing data and a HTTP request to Blueberry datastore.
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

        /// <summary>
        /// getFeched() makes a HTTP request to a Blueberry cloud and returns a list of all
        /// BAPI data structures which have been previously fetched in the current workbook.
        /// </summary>
        /// <returns></returns>
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
