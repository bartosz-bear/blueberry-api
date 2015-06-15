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
using Excel = Microsoft.Office.Interop.Excel;
using PublishingHelpers = ExcelAddIn1.Controllers.Helpers.PublishingHelpers;
using BlueberryRibbon = ExcelAddIn1.BlueberryRibbon;

namespace ExcelAddIn1.Controllers
{
    class Publishing
    {

        /// <summary>
        /// Gets data from Excel spreadsheet and send it to the BlueberryAPI cloud for persistent storage.
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

                xlRange = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                xlPath = xlWorkBook.Path;
                xlWorkbookName = xlWorkBook.Name;
                xlWorksheetName = xlWorkSheet.Name;
                xlDestinationCell = xlRange.Cells[1, 1].Address;
                xlName = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingNameTextBox.Text;
                xlDescription = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingDescriptionTextBox.Text;
                xlOrganization = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingOrganizationTextBox.Text;
                xlDataOwner = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingDataOwnerTextBox.Text;
                dataFromExcel = PublishingHelpers.measureData(xlRange, "noType");
                xlType = dataFromExcel["data_type"];
                xlID = xlOrganization.Replace(" ", "_") + "." + xlName.Replace(" ", "_") + "." + xlType;

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
                xlRange = PublishingHelpers.specifyRange(xlWorkSheet, xlDestinationCell, xlType);
                xlID = singleResult["bapi_id"];

                dataFromExcel = PublishingHelpers.measureData(xlRange, xlType);
                dataFromExcel["data_type"] = xlType;
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


            var httpWebRequest = (HttpWebRequest)WebRequest.Create(BlueberryRibbon.blueberryAPIurl + xlType + ".publish");
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
            var json = jsonSerializer.Serialize(activeWorkbookInfo);

            String url = BlueberryRibbon.blueberryAPIurl + "Data.get_published";

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

    }
}
