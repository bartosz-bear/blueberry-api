using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAddIn1.Utils;
using PublishingValidators = ExcelAddIn1.Controllers.Validators.PublishingValidators;
using System.Collections;
using System.Collections.Specialized;

namespace ExcelAddIn1.Controllers.Helpers
{
    class Helpers
    {
        public static bool validateRanges(Dictionary<string, dynamic> publishedData)
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

        public static Dictionary<string, dynamic> selectValidWorksheets(Dictionary<string, dynamic> publishedData)
        {
            ArrayList publishedSheets = publishedData["worksheets"];
            Microsoft.Office.Interop.Excel.Sheets worksheets = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets;
            ArrayList worksheetsArray = new ArrayList() { };
            foreach (Excel.Worksheet w in worksheets)
            {
                worksheetsArray.Add(w.Name);
            }
            foreach (string publishedSheet in publishedSheets.ToArray())
            {
                if (!worksheetsArray.Contains(publishedSheet))
                {
                    publishedData = removeMissingSheets(publishedData, publishedSheet);
                }
            }
            return publishedData;
        }

        private static Dictionary<string, dynamic> removeMissingSheets(Dictionary<string, dynamic> publishedData, string missingSheet)
        {

            List<string> keys = new List<string>();
            foreach (KeyValuePair<string, dynamic> kvp in publishedData)
            {
                keys.Add(kvp.Key);
            }
            List<int> itemsToRemove = new List<int>();
            for (int i = 0; i < publishedData["worksheets"].Count; i++)
            {
                if (publishedData["worksheets"][i] == missingSheet)
                {
                    itemsToRemove.Add(i);
                }
            }

            OrderedDictionary orderedDictionary = (OrderedDictionary)Utils.Utils.toOrderedDictionary(publishedData);

            for (int j = 0; j < itemsToRemove.Count; j++)
            {
                for (int k = 0; k < keys.Count; k++)
                {
                    ArrayList tempList = (ArrayList)orderedDictionary[keys[k]];
                    tempList.RemoveAt(itemsToRemove[j]);
                    orderedDictionary[keys[k]] = tempList;
                }
            }

            return (Dictionary<string, dynamic>)Utils.Utils.toDictionary(orderedDictionary);

        }

    }
}
