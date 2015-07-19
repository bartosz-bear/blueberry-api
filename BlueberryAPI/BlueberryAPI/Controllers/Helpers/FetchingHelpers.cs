using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Controllers.Helpers
{
    /// <summary>
    /// FetchingHelpers is a class container for all supporting methods necessary in 'Fetching' class.
    /// </summary>
    class FetchingHelpers
    {
        /// <summary>
        /// saveToExcel() method is a high level method which takes a single, serialized string containing
        /// description and actual BAPI data structures which needs to be saved to Excel. According to the actual
        /// data typeThis method calls another
        /// method saveArrayToExcel() which is responsible only for the actual load to excel.
        /// </summary>
        /// <param name="result"></param>
        public static void saveToExcel(Dictionary<string, dynamic> fetchedData)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlStartRange;
            Excel.Range xlEndRange;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            var jsonSerializer = new JavaScriptSerializer();
            String[] splitWords = fetchedData["bapi_id"].Split('.');
            string xlType = splitWords[2];
            string serializedData = fetchedData["data"][0];

            if (fetchedData["destination_cell"] != null)
            {
                xlStartRange = (Excel.Range)xlWorkSheet.Range[fetchedData["destination_cell"]];
            }
            else
            {
                xlStartRange = (Excel.Range)xlWorkSheet.Application.Selection;
            }

            //Depending on the BAPI data type, data will be saved in a different way.
            switch (xlType)
            {
                case "Scalar":
                    {
                        List<object> bapiData = new List<object>(); 
                        bapiData.Add(jsonSerializer.Deserialize<object>(serializedData));
                        //bapiData.Add(serializedData);
                        saveArrayToExcel(bapiData, xlWorkSheet, xlStartRange, 0);
                        break;
                    }
                case "List":
                    {
                        List<object> bapiData = jsonSerializer.Deserialize<List<object>>(serializedData);
                        saveArrayToExcel(bapiData, xlWorkSheet, xlStartRange, 0);
                        break;
                    }
                case "Dictionary":
                    {
                        List<List<object>> bapiData = jsonSerializer.Deserialize<List<List<object>>>(serializedData);
                        Int32 numOfListsInATable = bapiData.Count - 1;
                        for (var i = 0; i <= numOfListsInATable; i++)
                        {
                            List<object> currentBAPIDataArray = bapiData[i];
                            int offset = i;
                            saveArrayToExcel(currentBAPIDataArray, xlWorkSheet, xlStartRange, offset);
                        }

                        break;

                    }
                case "Table":
                    {
                        List<List<object>> bapiData = jsonSerializer.Deserialize<List<List<object>>>(serializedData);
                        Int32 numOfListsInATable = bapiData.Count - 1;
                        for (var i = 0; i <= numOfListsInATable; i++)
                        {
                            List<object> currentBAPIDataArray = bapiData[i];
                            int offset = i;
                            saveArrayToExcel(currentBAPIDataArray, xlWorkSheet, xlStartRange, offset);
                        }

                            break;
                    }
                default:
                    {
                        break;
                    }
            }

        }

        /// <summary>
        /// saveArrayToExcel saves an array of values to Excel range. In case offset is higher than zero, data
        /// will pasted in a range shifted to the 'offset' columns.
        /// </summary>
        /// <param name="arrayToBeSaved"></param>
        /// <param name="xlWorkSheet"></param>
        /// <param name="xlStartRange"></param>
        /// <param name="offset">Defines how many columns to the right should be range by shifted to.</param>
        public static void saveArrayToExcel(List<object> arrayToBeSaved, Excel.Worksheet xlWorkSheet, Excel.Range xlStartRange, Int32 offset)
        {
            Int32 dataLength = arrayToBeSaved.Count;

            xlStartRange = xlWorkSheet.Cells[xlStartRange.Row, xlStartRange.Column + offset];
            Excel.Range xlEndRange = (Excel.Range)xlWorkSheet.Cells[xlStartRange.Row + dataLength - 1, xlStartRange.Column];
            Excel.Range xlDestinationRange = xlWorkSheet.Range[xlStartRange, xlEndRange];

            var dataArray = new object[dataLength, 1];
            for (var i = 0; i < dataLength; i++)
            {
                dataArray[i, 0] = arrayToBeSaved[i];
            }

            xlDestinationRange.Value2 = dataArray;
        }

        /// <summary>
        /// It's a validation method which checks two Blueberry ID against two criterias: 1) Does the Blueberry ID has
        /// two dots. 2) Does a word after last dot is one of BAPI data types?
        /// </summary>
        /// <returns>Returns true if both criteria are satisfied.</returns>
        public static bool validateIDpreFetch()
        {
            string xlBlueberryID = Globals.Ribbons.Ribbon1.IDBox.Text;
            String[] splitWords;
            splitWords = xlBlueberryID.Split('.');
            
            // Validate whether the Blueberry ID has two dots
            try
            {
                var trying = splitWords[2];
            }
            catch (IndexOutOfRangeException ex)
            {
                MessageBox.Show("You have entered an incorrect Blueberry ID. Check the ID and try again.");
                return false;
            }

            // Validate whether a word after the last dot is one of BAPI data types defined in the acceptableDataTypes
            string xlDataType = splitWords[2];
            string[] acceptableDataTypes = { "Scalar", "List", "Dictionary", "Table" };

            if (!acceptableDataTypes.Contains(xlDataType))
            {
                MessageBox.Show("You have entered an incorrect Blueberry ID. Check the ID and try again.");
                return false;
            }
            return true;
        }

        /// <summary>
        /// It's a validation method to check whether the BAPI ID was found in Bluberry cloud datastore
        /// </summary>
        /// <param name="fetchedData">fetchedData represents a response from request to Blueberry cloud.
        /// If they key "info" returns "Incorrect BAPI ID" it indicates that the Blueberry ID was not found in Blueberry cloud.</param>
        /// <returns></returns>
        public static bool validateIDPostFetch(Dictionary<string, dynamic> fetchedData, string senderLabel)
        {
            // Validate whether the BAPI ID was found in Bluberry cloud datastore
            
            //var jsonSerializer = new JavaScriptSerializer();
            //Dictionary<string, dynamic> responseData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(fetchedData);


            if (fetchedData.ContainsKey("info"))
            {
                if (senderLabel == "Download")
                {
                    MessageBox.Show("You have entered an incorrect Blueberry ID. Check the ID and try again.");
                    return false;
                }

                else if (senderLabel == "Refresh")
                {
                    return false;
                }
            }

            return true;

            /*
            if (fetchedData.ContainsKey("info")) {
                if ((string)fetchedData["info"] == "Incorrect BAPI ID")
                {
                    MessageBox.Show("You have entered an incorrect Blueberry ID. Check the ID and try again.");
                    return false;
                }
                else if ((string)fetchedData["info"] == "One or more data points which you are trying to fetch doesn't exist anymore.")
                {
                    
                }
            }
            return true;
             */ 
        }

    }
}
