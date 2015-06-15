using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Net;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Controllers.Helpers
{
    class PublishingHelpers
    {

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
                    Dictionary<string, dynamic> publishingDictionary = new Dictionary<string, dynamic>();
                    for (int currentRowsCount = 1; currentRowsCount <= rowsCount; currentRowsCount++)
                    {
                        publishingDictionary.Add((dynamic)(xlRange.Cells[currentRowsCount, 1] as Excel.Range).Value2,
                                                 (dynamic)(xlRange.Cells[currentRowsCount, 2] as Excel.Range).Value2);
                    }
                    var jsonDictSerializer = new JavaScriptSerializer();
                    var jsonDict = jsonDictSerializer.Serialize(publishingDictionary);

                    return jsonDict;
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
                    var jsonTableSerializer = new JavaScriptSerializer();
                    var jsonTable = jsonTableSerializer.Serialize(publishingTable);

                    return jsonTable;
                default:
                    return "Other";
            }
        }

        public static Dictionary<string, dynamic> measureData(Excel.Range xlRange, String xlDataType)
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

        public static Excel.Range specifyRange(Excel.Worksheet xlWorkSheet, string xlDestinationCell, string xlType)
        {
            Excel.Range xlStartRange;
            Excel.Range xlEndRange;
            Excel.Range xlRange;
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
            return xlRange;
        }

        public static Boolean isPublishRangeEmpty(Excel.Range xlRange)
        {
            if (xlRange.Value2 == null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

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

        public static Boolean isIDUsed(string xlName, string xlOrganization, string xlDataOwner, Excel.Range xlRange)
        {

            string xlID = xlOrganization.Replace(" ", "_") + "." + xlName.Replace(" ", "_") + "." + labelData(xlRange.Rows.Count, xlRange.Columns.Count);

            Dictionary<string, dynamic> requestData = new Dictionary<string, dynamic>();
            requestData.Add("bapi_id", xlID);
            requestData.Add("user", xlDataOwner);

            var jsonSerializer = new JavaScriptSerializer();
            var json = jsonSerializer.Serialize(requestData);


            var httpWebRequest = (HttpWebRequest)WebRequest.Create(BlueberryRibbon.blueberryAPIurl + "Data.is_id_used");
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

                //Dictionary<string, bool> validations = new Dictionary<string, bool>();
                Dictionary<string, bool> deserializedID = jsonSerializer.Deserialize<Dictionary<string, bool>>(result);

                return deserializedID["response"];
            }

        }

        public static Boolean areInputsSpecialCharactersFree(string xlName, string xlDescription, string xlOrganization, string xlDataOwner)
        {

            var regexItem = new Regex(@"^[\w\s-]{1,80}$");
            List<string> items = new List<string>();
            items.Add(xlName);
            items.Add(xlOrganization);

            foreach (string i in items)
            {
                if (regexItem.IsMatch(i)) {
                    continue;
                }
                else
                {
                    return true;
                }
            }
            return false;

        }

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
            string xlDataOwner = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingDataOwnerTextBox.Text;

            Dictionary<string, bool> validations = new Dictionary<string, bool>();
            Dictionary<string, string> errorMessages = new Dictionary<string, string>();

            errorMessages.Add("isPublishRangeEmpty", "Range which you are trying to publish is empty. Choose some data and try again.");
            errorMessages.Add("isAnyBlueberryTaskPaneFieldEmpty", "One of the input forms ('Name', 'Description', 'Organization', 'Data Owner')" +
                                                                    " is empty. Please complete all fields before submitting.");
            errorMessages.Add("isIDUsed", "This 'Name' has already been used within this 'Organization' by a different user. Please change one or both of" +
                                          " them and try again.");
            errorMessages.Add("areInputsSpecialCharactersFree", "'Name' and 'Organization' should not have any of the following characters: '/*-+@&$#%.,\\\"'" +
                                                                " and it should be less than 80 characters.");

            validations.Add("isPublishRangeEmpty", isPublishRangeEmpty(xlRange));
            validations.Add("isAnyBlueberryTaskPaneFieldEmpty", isAnyBlueberryTaskPaneFieldEmpty(xlName, xlDescription, xlOrganization, xlDataOwner));
            validations.Add("isIDUsed", isIDUsed(xlName, xlOrganization, xlDataOwner, xlRange));
            validations.Add("areInputsSpecialCharactersFree", areInputsSpecialCharactersFree(xlName, xlDescription, xlOrganization, xlDataOwner));

            foreach (KeyValuePair<string, bool> item in validations)
            {
                if (item.Value)
                {
                    string itemKey = item.Key;
                    return errorMessages[itemKey];
                }
                    
            }
            return "Pass";
        }

    }
}
