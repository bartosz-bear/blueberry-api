using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAddIn1.Utils;
using System.Collections.Specialized;
using PublishingHelpers = ExcelAddIn1.Controllers.Helpers.PublishingHelpers;
using System.Web.Script.Serialization;
using System.Net;
using System.IO;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Reflection;

namespace ExcelAddIn1.Controllers.Validators
{
    class PublishingValidators
    {
            //private Excel.Workbook xlWorkBook;
            //private Excel.Worksheet xlWorkSheet;
            //private Excel.Range xlRange;
            //private string xlName;
            //private string xlDescription;
            //private string xlOrganization;
            //private string xlDataOwner;
            private OrderedDictionary errorMessages;
            private OrderedDictionary validatorsArguments;

            public OrderedDictionary ValidatorsArguments
            {
                get { return validatorsArguments; }
                set { validatorsArguments = value; }
            }

            public PublishingValidators(Excel.Range xlRange = null,  string xlName = null, string xlDescription = null,
                                        string xlOrganization = null, string xlDataOwner = null)
            {
                //this.xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
                //this.xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;
                //this.xlRange = (Excel.Range)xlWorkSheet.Application.Selection;
                this.errorMessages = new OrderedDictionary();
                validatorsArguments = new OrderedDictionary();
                this.errorMessages.Add("isPublishRangeEmpty", "Range which you are trying to publish is empty. Choose some data and try again.");
                this.errorMessages.Add("isAnyBlueberryTaskPaneFieldEmpty", "One of the input forms ('Name', 'Description', 'Organization', 'Data Owner')" +
                                                                        " is empty. Please complete all fields before submitting.");
                this.errorMessages.Add("isIDUsed", "This 'Name' has already been used within this 'Organization' by a different user. Please change one or both of" +
                                              " them and try again.");
                this.errorMessages.Add("areInputsSpecialCharactersFree", "'Name' and 'Organization' should not have any of the following characters: '/*-+@&$#%.,\\\"'" +
                                                                    " and it should be less than 80 characters.");
                validatorsArguments.Add("isPublishRangeEmpty", new object[1] { xlRange });
                validatorsArguments.Add("isAnyBlueberryTaskPaneFieldEmpty", new object[4] { xlName, xlDescription, xlOrganization, xlDataOwner });
                validatorsArguments.Add("isIDUsed", new object[4] { xlName, xlOrganization, xlDataOwner, xlRange });
                validatorsArguments.Add("areInputsSpecialCharactersFree", new object[4] { xlName, xlDescription, xlOrganization, xlDataOwner });
            }

            public string validatePublishingInputs(List<string> validators) {

                foreach (string v in validators)
                {
                    object[] validatorArguments = (object[])ValidatorsArguments[v];
                    bool returnedFalse = (bool)this.GetType().GetMethod(v, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(this, validatorArguments);
                    if (returnedFalse) { return (string)this.errorMessages[v]; }
                }
                return "Pass";
            }

            /// <summary>
            /// It's a validation method to make sure that the user is not publishing an empty data.
            /// </summary>
            /// <param name="xlRange"></param>
            /// <returns></returns>
            private Boolean isPublishRangeEmpty(Excel.Range xlRange)
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

            /*
            public Boolean isPublishRangeEmpty(Excel.Range xlRange)
            {
                string errorMessage = "Range which you are trying to publish is empty. Choose some data and try again.";
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
             */ 

            /// <summary>
            /// It's a validation method to make sure that all fields in the BlueberryTaskPane are not-empty. It checks 'Name',
            /// 'Description', 'Organization' and 'Data owner'.
            /// </summary>
            /// <param name="xlName"></param>
            /// <param name="xlDescription"></param>
            /// <param name="xlOrganization"></param>
            /// <param name="xlDataOwner"></param>
            /// <returns></returns>
            private Boolean isAnyBlueberryTaskPaneFieldEmpty(string xlName, string xlDescription, string xlOrganization, string xlDataOwner)
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
            private Boolean isIDUsed(string xlName, string xlOrganization, string xlDataOwner, Excel.Range xlRange)
            {
                
                
                string xlID = xlOrganization.Replace(" ", "_") + "." + xlName.Replace(" ", "_") + "." + PublishingHelpers.labelData(xlRange.Rows.Count, xlRange.Columns.Count);

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

            private dynamic isIDUsedHandleResponse(object[] args)
            {
                var serializer = new JavaScriptSerializer();
                StreamReader streamReader = (StreamReader)args[0];
                string result = streamReader.ReadToEnd();
                Dictionary<string, bool> isIDUsedResponse = serializer.Deserialize<Dictionary<string, bool>>(result);
                return isIDUsedResponse["response"];
            }

            private dynamic isIDUsedHandleExceptions(object[] args)
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
            private Boolean areInputsSpecialCharactersFree(string xlName, string xlDescription, string xlOrganization, string xlDataOwner)
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
    }
}
