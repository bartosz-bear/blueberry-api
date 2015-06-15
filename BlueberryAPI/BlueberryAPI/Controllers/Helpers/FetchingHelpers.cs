using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1.Controllers.Helpers
{
    class FetchingHelpers
    {
        public static void saveToExcel(string result)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlStartRange;
            Excel.Range xlEndRange;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            var jsonSerializer = new JavaScriptSerializer();
            Dictionary<string, dynamic> fetchedData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(result);
            String[] splitWords = fetchedData["bapi_id"].Split('.');
            string xlType = splitWords[2];
            string serializedData = fetchedData["data"][0];

            if (fetchedData["destination_cell"] != null)
            {
                xlStartRange = (Excel.Range)xlWorkSheet.Range[fetchedData["destination_cell"]];
            }
            else
            {
                //Make sure that this xlStartRange is a single range. If the user selects more than a single range, then choose the top one.
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
                        Dictionary<string, dynamic> bapiData = jsonSerializer.Deserialize<Dictionary<string, dynamic>>(serializedData);

                        var bapiDataKeys = bapiData.Keys.ToArray();
                        var bapiDataValues = bapiData.Values.ToArray();
                        Int32 dataLength = bapiData.Keys.Count;

                        // Saving keys to excel
                        // TODO: Refactor this to use saveArrayToExcel 
                        Excel.Range xlEndRangeForKeys = (Excel.Range)xlWorkSheet.Cells[xlStartRange.Row + dataLength - 1, xlStartRange.Column];
                        Excel.Range xlDestinationRangeForKeys = xlWorkSheet.Range[xlStartRange, xlEndRangeForKeys];

                        var keysDataArray = new object[dataLength, 1];
                        for (var i = 0; i < dataLength; i++)
                        {
                            keysDataArray[i, 0] = bapiDataKeys[i];
                        }

                        xlDestinationRangeForKeys.Value2 = keysDataArray;

                        // Saving values to excel
                        xlStartRange = (Excel.Range)xlWorkSheet.Cells[xlStartRange.Row, xlStartRange.Column + 1];
                        Excel.Range xlEndRangeForValues = (Excel.Range)xlWorkSheet.Cells[xlStartRange.Row + dataLength - 1, xlStartRange.Column];
                        Excel.Range xlDestinationRangeForValues = xlWorkSheet.Range[xlStartRange, xlEndRangeForValues];

                        var valuesDataArray = new object[dataLength, 1];
                        for (var i = 0; i < dataLength; i++)
                        {
                            valuesDataArray[i, 0] = bapiDataValues[i];
                        }

                        xlDestinationRangeForValues.Value2 = valuesDataArray;

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
                        /*
                        Int32 dataLength = fetchedData["data"].Count;
                        xlEndRange = (Excel.Range)xlWorkSheet.Cells[xlStartRange.Row + dataLength - 1, xlStartRange.Column];
                        Excel.Range xlDestinationRange = xlWorkSheet.Range[xlStartRange, xlEndRange];

                        var fetchedDataArray = new object[dataLength, 1];
                        for (var i = 0; i < dataLength; i++)
                        {
                            fetchedDataArray[i, 0] = fetchedData["data"][i];
                        }

                        xlDestinationRange.Value2 = fetchedDataArray;
                        break;
                         */
                        break;
                    }
            }

        }

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
    }
}
