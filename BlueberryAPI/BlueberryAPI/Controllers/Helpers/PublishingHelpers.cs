using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Net;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Ciloci.Flee;
using System.Windows.Forms;
using ExcelAddIn1.Utils;
using PublishingValidators = ExcelAddIn1.Controllers.Validators.PublishingValidators;

namespace ExcelAddIn1.Controllers.Helpers
{
    /// <summary>
    /// PublishingHelpers is a class container for all supporting methods necessary in 'Publishing' class.
    /// </summary>
    class PublishingHelpers
    {
        /// <summary>
        /// fromExcelToObject() method takes an Excel range, fetches data from Excel, converts it to C# object and serialize it to JSON.
        /// There are four data types which can be fetched from Excel and converted to JSON objects:  Scalar (one Excel cell), List (several
        /// Excel cells in one column), Dictionary (several Excel cells in two columns, typically the first column represents keys and
        /// the second column represents values), Table (several Excel cells in several columns).
        /// </summary>
        /// <param name="rowsCount"></param>
        /// <param name="columnsCount"></param>
        /// <param name="dataType"></param>
        /// <param name="xlRange"></param>
        /// <returns>Returns serialized JSON object which represents an Excel range.</returns>
        public static dynamic fromExcelToObject(int rowsCount, int columnsCount, string dataType, Excel.Range xlRange)
        {
            switch (dataType)
            {
                case "Scalar":
                    var jsonScalarSerializer = new JavaScriptSerializer();
                    var jsonScalar = jsonScalarSerializer.Serialize(xlRange.Value2);
                    return jsonScalar;

                case "List":
                    ArrayList publishingList = new ArrayList();
                    for (int currentRowsCount = 1; currentRowsCount <= rowsCount; currentRowsCount++)
                    {
                        for (int currentColumnsCount = 1; currentColumnsCount <= columnsCount; currentColumnsCount++)
                        {
                            publishingList.Add((dynamic)(xlRange.Cells[currentRowsCount, currentColumnsCount] as Excel.Range).Value2);
                        }
                    }
                    var jsonListSerializer = new JavaScriptSerializer();
                    var jsonList = jsonListSerializer.Serialize(publishingList);
                    return jsonList;

                case "Dictionary":
                    List<List<object>> publishingDictionary = new List<List<object>>();
                    int columnsCountCopyForDict = columnsCount;
                    for (int currentColumnsCount = 1; currentColumnsCount <= columnsCount; currentColumnsCount++)
                    {
                        List<object> sublist = new List<object>();
                        for (int currentRowsCount = 1; currentRowsCount <= rowsCount; currentRowsCount++)
                        {
                            sublist.Add((dynamic)(xlRange.Cells[currentRowsCount, currentColumnsCount] as Excel.Range).Value2);
                        }
                        publishingDictionary.Add(sublist);
                    }

                    var jsonDictSerializer = new JavaScriptSerializer();
                    var jsonDict = jsonDictSerializer.Serialize(publishingDictionary);
                    return jsonDict;

                case "Table":
                    List<List<object>> publishingTable = new List<List<object>>();
                    int columnsCountCopy = columnsCount;
                    for (int currentColumnsCount = 1; currentColumnsCount <= columnsCount; currentColumnsCount++)
                    {
                        List<object> sublist = new List<object>();
                        for (int currentRowsCount = 1; currentRowsCount <= rowsCount; currentRowsCount++)
                        {
                            sublist.Add((dynamic)(xlRange.Cells[currentRowsCount, currentColumnsCount] as Excel.Range).Value2);
                        }
                        publishingTable.Add(sublist);
                    }
                    var jsonTableSerializer = new JavaScriptSerializer();
                    var jsonTable = jsonTableSerializer.Serialize(publishingTable);
                    return jsonTable;

                default:
                    return "Other";
            }
        }

        /// <summary>
        /// measureData() creates a dictionary which serves as a meta-data about the Range which was 
        /// passed as an argument. It also contains data fetched from Excel in a serialized JSON format.
        /// </summary>
        /// <param name="xlRange">Excel Range to be described and measured.</param>
        /// <param name="xlDataType">xlDataType is either one of the BAPI data types (eg.Scalar, List, etc)
        /// or the argument is not passed at all. In the second case the default argument is is "noType" which
        /// means that the data type should be infered from the size of the Range using labelData() method.</param>
        /// <returns> The dictionary provides information 1) number of rows, 2) number of columns
        /// 3) type of data (Scalar, List etc). The last key points to a JSON serialized object fetched from
        /// the Excel range which was passed as argument.
        /// </returns>
        public static Dictionary<string, dynamic> measureData(Excel.Range xlRange, String xlDataType = "noType")
        {
            int xlRowsCount;
            int xlColumnsCount;
            xlRowsCount = xlRange.Rows.Count;
            xlColumnsCount = xlRange.Columns.Count;

            Dictionary<string, dynamic> dataInfo = new Dictionary<string, dynamic>();
            dataInfo.Add("rows_count", xlRowsCount);
            dataInfo.Add("columns_count", xlColumnsCount);
            if (xlDataType == "noType")
            {
                dataInfo.Add("data_type", labelData(xlRowsCount, xlColumnsCount));
            }
            else
            {
                dataInfo.Add("data_type", xlDataType);
            }

            dataInfo.Add("data", fromExcelToObject(xlRowsCount, xlColumnsCount, dataInfo["data_type"], xlRange));
            return dataInfo;
        }

        /// <summary>
        /// labelData() returns one of the BAPI data types. The type depends on the number or rows and columns
        /// of the Range passed.
        /// </summary>
        /// <param name="xlRowsCount"></param>
        /// <param name="xlColumnsCount"></param>
        /// <returns>Returns a string indicating what BAPI data type the passed Range is.</returns>
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
                    if (xlColumnsCount > 2)
                    {
                        return "Table";
                    }
                    else
                    {
                        return "Other";
                    }
            }
        }

        /// <summary>
        /// specifyRange() calculates an Excel Range where the top left cell would be the initial destination cell - 
        /// xlDestinationCell. The bottom right cell of the Range depends on the BAPI data type.
        /// </summary>
        /// <param name="xlWorkSheet"></param>
        /// <param name="xlDestinationCell"></param>
        /// <param name="xlType"></param>
        /// <returns>Returns a new Excel Range indicating where data should be fetched from or saved to.</returns>
        public static Excel.Range specifyRange(Excel.Worksheet xlWorkSheet, string xlDestinationCell, string xlType)
        {
            Excel.Range xlStartRange;
            Excel.Range xlEndRange;
            Excel.Range xlRange;

            xlRange = (Excel.Range)xlWorkSheet.Range[xlDestinationCell];
            
            // THIS FUNCTION WILL BE RESPONSIBLE TO UPDATING DATA ACCORDING TO USER CHOICE - EXTEND DATA STRUCTURE OR NOT
            
            /*
            xlStartRange = (Excel.Range)xlWorkSheet.Range[xlDestinationCell];

            switch (xlType)
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
            */
            return xlRange;
        }

        /// <summary>
        /// It's a validation method to make sure that the user is not publishing an empty data.
        /// </summary>
        /// <param name="xlRange"></param>
        /// <returns></returns>
        public static Boolean isPublishRangeEmpty(Excel.Range xlRange)
        {
            switch (xlRange.Count)
            {
                case 1:
                    {
                        if (xlRange.Value2 == null) { return true; };
                        return false;
                    }
                default:
                    {
                        foreach (var cellValue in xlRange.Value2)
                        {
                            if (cellValue != null) { return false; }
                        }
                        return true;
                    }
            }
        }

        /// <summary>
        /// It's a validation method to make sure that all fields in the BlueberryTaskPane are not-empty. It checks 'Name',
        /// 'Description', 'Organization' and 'Data owner'.
        /// </summary>
        /// <param name="xlName"></param>
        /// <param name="xlDescription"></param>
        /// <param name="xlOrganization"></param>
        /// <param name="xlDataOwner"></param>
        /// <returns></returns>
        public static Boolean isAnyBlueberryTaskPaneFieldEmpty(string xlName, string xlDescription, string xlOrganization, string xlDataOwner)
        {

            if (string.IsNullOrEmpty(xlName) || string.IsNullOrEmpty(xlDescription) ||
                string.IsNullOrEmpty(xlOrganization) || string.IsNullOrEmpty(xlDataOwner))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// It's a validation method to check whether a particular combination of a 'Name', 'Organization' and BAPI data type
        /// has not been used before. The combination of these three parameters is used to create a BAPI ID. Using existing BAPI ID
        /// by the same 'Data Owner' is acceptable. Using existing BAPI ID created by a different 'Data Owner' is not acceptable.
        /// In order to find out which BAPI ID has already been used this method sends a HTTP request to the Blueberry API /Data.is_id_used.
        /// </summary>
        /// <param name="xlName"></param>
        /// <param name="xlOrganization"></param>
        /// <param name="xlDataOwner"></param>
        /// <param name="xlRange"></param>
        /// <returns>If BAPI ID is existing and was used by a different 'Data Ownder' returns 'True'. Otherwise returns 'False'.</returns>
        public static Boolean isIDUsed(string xlName, string xlOrganization, string xlDataOwner, Excel.Range xlRange)
        {

            string xlID = xlOrganization.Replace(" ", "_") + "." + xlName.Replace(" ", "_") + "." + labelData(xlRange.Rows.Count, xlRange.Columns.Count);

            Dictionary<string, dynamic> requestData = new Dictionary<string, dynamic>();
            requestData.Add("bapi_id", xlID);
            requestData.Add("user", xlDataOwner);

            var jsonSerializer = new JavaScriptSerializer();
            var data = jsonSerializer.Serialize(requestData);

            var httpWebRequest = (HttpWebRequest)WebRequest.Create(GlobalVariables.blueberryAPIurl + "Data.is_id_used");
            httpWebRequest.ContentType = "text/json";
            httpWebRequest.Method = "POST";

            // Send an HTTP request to Blueberry Cloud to verify whether the ID has been used before.
            object[] httpResponseArgs = new object[] { "StreamReaderProperty" };
            BlueberryHTTPResponse httpResponse = new BlueberryHTTPResponse(httpWebRequest, data, httpResponseArgs);

            return (bool)httpResponse.sendHTTPRequest(new BlueberryHTTPResponse.handleResponseDelegate(isIDUsedHandleResponse),
                new BlueberryHTTPResponse.handleReponseExceptionsDelegate(isIDUsedHandleExceptions));
        }

        private static dynamic isIDUsedHandleResponse(object[] args)
        {
            var serializer = new JavaScriptSerializer();
            StreamReader streamReader = (StreamReader)args[0];
            string result = streamReader.ReadToEnd();
            Dictionary<string, bool> isIDUsedResponse = serializer.Deserialize<Dictionary<string, bool>>(result);
            return isIDUsedResponse["response"];
        }

        private static dynamic isIDUsedHandleExceptions(object[] args)
        {
            MessageBox.Show("Please connect to Internet.");
            return false;
        }

        /// <summary>
        /// It's a validation method which checks that none of the following characters '/*-+@&$#%.,\" have
        /// been used in any of the BlueberryTaskBane fields.
        /// </summary>
        /// <param name="xlName"></param>
        /// <param name="xlDescription"></param>
        /// <param name="xlOrganization"></param>
        /// <param name="xlDataOwner"></param>
        /// <returns>If all of the fields are free of all of the special characters the method returns 'False'.</returns>
        public static Boolean areInputsSpecialCharactersFree(string xlName, string xlDescription, string xlOrganization, string xlDataOwner)
        {

            var regexItem = new Regex(@"^[\w\s-]{1,80}$");
            List<string> items = new List<string>();
            items.Add(xlName);
            items.Add(xlOrganization);

            foreach (string i in items)
            {
                if (regexItem.IsMatch(i))
                {
                    continue;
                }
                else
                {
                    return true;
                }
            }
            return false;

        }

        /// <summary>
        /// It's a validation methods handler which runs all validation methods defined inside the method.
        /// </summary>
        /// <returns></returns>
        public static String validatePublishingInputs()
        {

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            xlRange = (Excel.Range)xlWorkSheet.Application.Selection;

            string xlName = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingNameTextBox.Text;
            string xlDescription = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingDescriptionTextBox.Text;
            string xlOrganization = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingOrganizationTextBox.Text;
            string xlDataOwner = GlobalVariables.sessionData["loggedUser"];

            OrderedDictionary errorMessages = new OrderedDictionary();

            errorMessages.Add("isPublishRangeEmpty", "Range which you are trying to publish is empty. Choose some data and try again.");
            errorMessages.Add("isAnyBlueberryTaskPaneFieldEmpty", "One of the input forms ('Name', 'Description', 'Organization', 'Data Owner')" +
                                                                    " is empty. Please complete all fields before submitting.");
            errorMessages.Add("areInputsSpecialCharactersFree", "'Name' and 'Organization' should not have any of the following characters: '/*-+@&$#%.,\\\"'" +
                                                                " and it should be less than 80 characters.");
            errorMessages.Add("isIDUsed", "This 'Name' has already been used within this 'Organization' by a different user. Please change one or both of" +
                                          " them and try again.");

            if (isPublishRangeEmpty(xlRange))
            {
                return (string)errorMessages["isPublishRangeEmpty"];
            }
            if (isAnyBlueberryTaskPaneFieldEmpty(xlName, xlDescription, xlOrganization, xlDataOwner))
            {
                return (string)errorMessages["isAnyBlueberryTaskPaneFieldEmpty"];
            }
            if (areInputsSpecialCharactersFree(xlName, xlDescription, xlOrganization, xlDataOwner))
            {
                return (string)errorMessages["areInputsSpecialCharactersFree"];
            }
            if (isIDUsed(xlName, xlOrganization, xlDataOwner, xlRange))
            {
                return (string)errorMessages["isIDUsed"];
            }

            return "Pass";
        }

        public static bool blueberryTaskPaneExists()
        {
            string currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            foreach (var i in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (i.Title == "Publish" + currentWorkbook)
                {
                    return true;
                }
            }
            return false;
        }

        public static bool blueberryTaskPaneVisible()
        {
            string currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            foreach (var i in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (i.Title == "Publish" + currentWorkbook && i.Visible)
                {
                    return true;
                }
            }
            return false;
        }

        public static void showBlueberryTaskPane()
        {
            string currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            foreach (var i in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (i.Title == "Publish" + currentWorkbook && !i.Visible)
                {
                    i.Visible = true;
                }
            }
        }

        public static bool validateUpdateRanges(Dictionary<string, dynamic> publishedData)
        {
            Excel.Workbook xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            int publishedDataItemsCount = publishedData["ids"].Count;
            for (int i = 0; i < publishedDataItemsCount; i++)
            {
                Dictionary<string, dynamic> singleResult = new Dictionary<string, dynamic>();
                var a = publishedData["worksheets"][i];
                Excel.Range xlRange = PublishingHelpers.specifyRange(xlWorkSheet, (string)publishedData["destination_cells"][i], (string)publishedData["data_types"][i]);
                PublishingValidators validator = new PublishingValidators(xlRange);
                string validationResult = validator.validatePublishingInputs(new List<string> { "isPublishRangeEmpty" });
                if (validationResult != "Pass") { return false; }
            }
            return true;
        }

        public static string publishSeveral(Dictionary<string, dynamic> publishedData)
        {
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
                Publishing.publishData(singleResult);
            }
            return "Data has been published.";
        }
    }
}
