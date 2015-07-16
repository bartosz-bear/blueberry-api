using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Net;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using PublishingHelpers = ExcelAddIn1.Controllers.Helpers.PublishingHelpers;
using PublishingValidators = ExcelAddIn1.Controllers.Validators.PublishingValidators;
using BlueberryRibbon = ExcelAddIn1.BlueberryRibbon;
using ExcelAddIn1.Utils;

namespace ExcelAddIn1.Controllers
{
    /// <summary>
    /// Publishing class it the container for methods responsible for publishing data to BlueberryAPI cloud.
    /// </summary>
    class Publishing
    {

        /// <summary>
        /// Gets data from Excel spreadsheet and send it to the Blueberry cloud for persistent storage.
        /// It works both for newly published data as activated by 'Publish' button on the Excel Add-in and
        /// for 'Update' button which first gets a list of previously published data and updates them with
        /// 'fresh' data.
        /// </summary>
        /// <param name="singleResult">singleResults helps to distinguish whether data is published 
        /// for the first time or it has already been published. If singleResult is null then data is
        /// published for the first time.
        /// </param>
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
                // Publishing data for the first time. Arguments are populated from BlueberryTaskPane.
                xlRange = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                xlPath = xlWorkBook.Path;
                xlWorkbookName = xlWorkBook.Name;
                xlWorksheetName = xlWorkSheet.Name;
                xlDestinationCell = xlRange.Cells.Address;
                xlName = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingNameTextBox.Text;
                xlDescription = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingDescriptionTextBox.Text;
                xlOrganization = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingOrganizationTextBox.Text;
                xlDataOwner = GlobalVariables.sessionData["loggedUser"];
                dataFromExcel = PublishingHelpers.measureData(xlRange);
                xlType = dataFromExcel["data_type"];
                xlID = xlOrganization.Replace(" ", "_") + "." + xlName.Replace(" ", "_") + "." + xlType;

            }
            else
            {
                // Data has been published before. Arguments not related to data are populated from PublishConfigurations
                // class which is stored in the BlueberryAPI cloud.
                xlPath = singleResult["workbook_path"];
                xlWorkbookName = singleResult["workbook"];
                xlWorksheetName = singleResult["worksheet"];
                xlDestinationCell = singleResult["destination_cell"];
                xlType = singleResult["data_type"];
                xlName = singleResult["name"];
                xlDescription = singleResult["description"];
                xlOrganization = singleResult["organization"];
                xlDataOwner = singleResult["user"];

                // Use specifyRange() and measureData() methods to retrieve the full Excel Range of data. 
                xlRange = PublishingHelpers.specifyRange(xlWorkSheet, xlDestinationCell, xlType);
                xlID = singleResult["bapi_id"];
                dataFromExcel = PublishingHelpers.measureData(xlRange, xlType);
                dataFromExcel["data_type"] = xlType;
            }

            // Collect all arguments in a Dictionary to be easily converted to JSON format
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

            // Serialize and send data via HTTP POST request
            var jsonSerializer = new JavaScriptSerializer();
            var data = jsonSerializer.Serialize(publishingData);

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(GlobalVariables.blueberryAPIurl + xlType + ".publish");
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            object[] httpResponseArgs = new object[] { "StreamReaderProperty" };
            BlueberryHTTPResponse httpResponse = new BlueberryHTTPResponse(httpWebRequest, data, httpResponseArgs);

            dynamic response = httpResponse.sendHTTPRequest(new BlueberryHTTPResponse.handleResponseDelegate(publishDataHandleResponse),
                new BlueberryHTTPResponse.handleReponseExceptionsDelegate(publishDataReturnResponse));

        }

        private static dynamic publishDataHandleResponse(object[] args)
        {
            var serializer = new JavaScriptSerializer();
            StreamReader streamReader = (StreamReader)args[0];
            string result = streamReader.ReadToEnd();
            Dictionary<string, dynamic> deserializedResult = serializer.Deserialize<Dictionary<string, dynamic>>(result);
            //MessageBox.Show(deserializedResult["response"]);
            return false;
        }

        private static dynamic publishDataReturnResponse(object[] args)
        {
            return true;
        }

        /// <summary>
        /// Send a request to Blueberry cloud in order to retrieve a list of previously published data.
        /// It retrieves all previously published data in the current workbook.
        /// </summary>
        /// <returns>Returns a dictionary with information such as ID, name, description</returns>
        public static Dictionary<string, dynamic> getPublished()
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
            var data = jsonSerializer.Serialize(activeWorkbookInfo);

            String url = GlobalVariables.blueberryAPIurl + "Data.get_published";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            object[] httpResponseArgs = new object[] { "StreamReaderProperty" };
            BlueberryHTTPResponse httpResponse = new BlueberryHTTPResponse(httpWebRequest, data, httpResponseArgs);

            return (dynamic)httpResponse.sendHTTPRequest(new BlueberryHTTPResponse.handleResponseDelegate(getPublishedHandleResponse),
                new BlueberryHTTPResponse.handleReponseExceptionsDelegate(getPublishedHandleExceptions));

        }

        private static dynamic getPublishedHandleResponse(object[] args)
        {
            var serializer = new JavaScriptSerializer();
            StreamReader streamReader = (StreamReader)args[0];
            string result = streamReader.ReadToEnd();
            Dictionary<string, dynamic> deserializedResult = serializer.Deserialize<Dictionary<string, dynamic>>(result);
            return deserializedResult;
        }

        private static dynamic getPublishedHandleExceptions(object[] args)
        {
            MessageBox.Show("Please connect to Internet.");
            return new Dictionary<string, dynamic>();
        }

    }
}
