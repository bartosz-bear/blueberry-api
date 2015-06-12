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
    class PublishingHelpers
    {

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
    }
}
