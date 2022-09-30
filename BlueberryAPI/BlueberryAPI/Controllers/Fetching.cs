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
using ExcelAddIn1.Utils;
using System.Collections;

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
        public static Dictionary<string, dynamic> fetchData(Dictionary<string, dynamic> singleResult = null)
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
                xlDataOwner = GlobalVariables.sessionData["loggedUser"];
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
            var data = jsonSerializer.Serialize(fetchingData);

            String[] splitWords = fetchingData["bapi_id"].Split('.');
            String url = GlobalVariables.blueberryAPIurl + splitWords[2] + ".fetch";
            
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            object[] httpResponseArgs = new object[] { "StreamReaderProperty" };
            BlueberryHTTPResponse httpResponse = new BlueberryHTTPResponse(httpWebRequest, data, httpResponseArgs);

            return httpResponse.sendHTTPRequest(new BlueberryHTTPResponse.handleResponseDelegate(fetchDataHandleResponse),
                new BlueberryHTTPResponse.handleReponseExceptionsDelegate(fetchDataHandleExceptions));

        }

        private static dynamic fetchDataHandleResponse(object[] args)
        {
            var serializer = new JavaScriptSerializer();
            StreamReader streamReader = (StreamReader)args[0];
            var jsonSerializer = new JavaScriptSerializer();
            Dictionary<string, dynamic> fetchedData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(streamReader.ReadToEnd());
            return fetchedData;
        }

        private static dynamic fetchDataHandleExceptions(object[] args)
        {
            switch ((string)args[0])
            {
                case "ProtocolError":
                    {
                        if ((string)args[1] == "The remote server returned an error: (500) Internal Server Error.")
                        {
                            MessageBox.Show("Blueberry ID doesn't exist.");
                        }
                        break;
                    }
                case "ConnectFailure":
                    {
                        MessageBox.Show("Please connect to Internet.");
                        break;
                    }
                case "NameResolutionFailure":
                    {
                        MessageBox.Show("Please connect to Internet.");
                        break;
                    }
                default:
                    {
                        MessageBox.Show("Something went wrong. Please investigate.");
                        break;
                    }
            }
            return new Dictionary<string, dynamic> { };
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
            var data = jsonSerializer.Serialize(activeWorkbookInfo);

            String url = GlobalVariables.blueberryAPIurl + "Data.get_fetched";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            object[] httpResponseArgs = new object[] { "StreamReaderProperty" };
            BlueberryHTTPResponse httpResponse = new BlueberryHTTPResponse(httpWebRequest, data, httpResponseArgs);

            return httpResponse.sendHTTPRequest(new BlueberryHTTPResponse.handleResponseDelegate(getFetchedHandleResponse),
                new BlueberryHTTPResponse.handleReponseExceptionsDelegate(getFetchedHandleExceptions));
        }

        private static dynamic getFetchedHandleResponse(object[] args)
        {
            var serializer = new JavaScriptSerializer();
            StreamReader streamReader = (StreamReader)args[0];
            string result = streamReader.ReadToEnd();
            return serializer.Deserialize<Dictionary<string, dynamic>>(result);
        }

        private static dynamic getFetchedHandleExceptions(object[] args)
        {
            MessageBox.Show("Please connect to Internet.");
            return new Dictionary<string, dynamic>();
        }

        public static List<ArrayList> getFavorites()
        {

            Dictionary<string, string> userInfo = new Dictionary<string, string>();
            userInfo.Add("user", GlobalVariables.sessionData["loggedUser"]);

            var jsonSerializer = new JavaScriptSerializer();
            var data = jsonSerializer.Serialize(userInfo);

            String url = GlobalVariables.blueberryAPIurl + "FavoriteID.get_favorite_ids";

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            object[] httpResponseArgs = new object[] { "StreamReaderProperty" };
            BlueberryHTTPResponse httpResponse = new BlueberryHTTPResponse(httpWebRequest, data, httpResponseArgs);

            var response = httpResponse.sendHTTPRequest(new BlueberryHTTPResponse.handleResponseDelegate(getFavoriteIDsHandleResponse),
                new BlueberryHTTPResponse.handleReponseExceptionsDelegate(getFavoriteIDsHandleExceptions));

            ArrayList favoriteIds = new ArrayList() { };
            ArrayList favoriteNames = new ArrayList() { };

            if (response.Count != 0)
            {
                favoriteIds = response["favorite_ids"];
                favoriteNames = response["favorite_names"];
            }

            return new List<ArrayList> {favoriteIds, favoriteNames};
        }

        private static dynamic getFavoriteIDsHandleResponse(object[] args)
        {
            var serializer = new JavaScriptSerializer();
            StreamReader streamReader = (StreamReader)args[0];
            string result = streamReader.ReadToEnd();
            return serializer.Deserialize<Dictionary<string, dynamic>>(result);
        }

        private static dynamic getFavoriteIDsHandleExceptions(object[] args)
        {
            return new Dictionary<string, dynamic>();
        }

        public static void createFavoriteButtons(ArrayList favoriteIDs, ArrayList favoriteNames, ExcelAddIn1.BlueberryRibbon currentRibbon)
        {

            var ribbonFactory = Globals.Factory.GetRibbonFactory();
            for (int i = 0; i < favoriteIDs.Count; i++)
            {
                RibbonButton tempButton = currentRibbon.Factory.CreateRibbonButton();
                tempButton.Description = (string)favoriteIDs[i];
                tempButton.Label = (string)favoriteNames[i];
                tempButton.Name = "FavoriteButton" + i.ToString();
                currentRibbon.FavoritesDropdown.Buttons.Add(tempButton);
                
            }
        }

        public static void populateFavoritesButtonsLabels(ArrayList favoriteIDs, ExcelAddIn1.BlueberryRibbon currentRibbon)
        {
            //string favoriteID = (string)favoriteIDs[0];
            //string[] favoriteIDSplit = favoriteID.Split('.');


            for (int i = 0; i < favoriteIDs.Count; i++)
            {
                currentRibbon.FavoritesDropdown.Buttons[i].Label = (string)favoriteIDs[i]; 
            }
            //currentRibbon.Fav1Button.Label = (string)favoriteIDs[0];
            //currentRibbon.Fav2Button.Label = (string)favoriteIDs[1];
            //currentRibbon.Fav3Button.Label = (string)favoriteIDs[2];
            //currentRibbon.Fav4Button.Label = (string)favoriteIDs[3];
            //currentRibbon.Fav5Button.Label = (string)favoriteIDs[4];
        }

        public static void populateFavoritesButtonsDescriptions(ArrayList favoriteNames, ExcelAddIn1.BlueberryRibbon currentRibbon)
        {
            for (int i = 0; i < favoriteNames.Count; i++)
            {
                currentRibbon.FavoritesDropdown.Buttons[i].Description = (string)favoriteNames[i];
            }
            
            //string favoriteID = (string)favoriteIDs[0];
            //string[] favoriteIDSplit = favoriteID.Split('.');
            //currentRibbon.Fav1Button.Description = (string)favoriteNames[0];
            //currentRibbon.Fav2Button.Description = (string)favoriteNames[1];
            //currentRibbon.Fav3Button.Description = (string)favoriteNames[2];
            //currentRibbon.Fav4Button.Description = (string)favoriteNames[3];
            //currentRibbon.Fav5Button.Description = (string)favoriteNames[4];
        }

        public static void hideUnusedFavoriteButtons(ArrayList favoriteIDs, ExcelAddIn1.BlueberryRibbon currentRibbon)
        {
            if (favoriteIDs.Count < 5)
            {
                int diff = 5 - favoriteIDs.Count;
                List<int> reverseIndices = new List<int>() { 4, 3, 2, 1, 0 };
                if (diff != 0)
                {
                    IEnumerable<int> range = Enumerable.Range(0, diff);
                    foreach (int num in range)
                    {
                        RibbonButton currentButton = currentRibbon.FavoritesDropdown.Buttons[reverseIndices[num]];
                        currentButton.Visible = false;
                    }
                }
            }
        }

        public static void showAllFavoriteButtons(ExcelAddIn1.BlueberryRibbon currentRibbon)
        {
            for (int i = 0; i < 5; i++)
            {
                RibbonButton currentButton = currentRibbon.FavoritesDropdown.Buttons[i];
                currentButton.Visible = true;
            }
        }

        public static void clearFavoritesButtons(int numberOfButtons, ExcelAddIn1.BlueberryRibbon currentRibbon)
        {
            for (int i = 0; i < numberOfButtons; i++)
            {
                currentRibbon.FavoritesDropdown.Buttons[i].Label = "";
            }
        }

    }
}
