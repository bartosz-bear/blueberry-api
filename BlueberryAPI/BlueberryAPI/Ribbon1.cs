using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Net;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using System.Data;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Publishing = ExcelAddIn1.Controllers.Publishing;
using Fetching = ExcelAddIn1.Controllers.Fetching;
using FetchingHelpers = ExcelAddIn1.Controllers.Helpers.FetchingHelpers;
using PublishingHelpers = ExcelAddIn1.Controllers.Helpers.PublishingHelpers;
using PublishingValidators = ExcelAddIn1.Controllers.Validators.PublishingValidators;
using Spring.Aop.Framework;
using ExcelAddIn1.Utils;
using System.Reflection;

namespace ExcelAddIn1
{
    public partial class BlueberryRibbon
    {
        public BlueberryTaskPane publishBlueberryTaskPane;
        public Microsoft.Office.Tools.CustomTaskPane myTaskPane;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
        }

        /* Buttons Events */

        private void Publish_Click(object sender, RibbonControlEventArgs e)
        {
            if (!UserManagement.userLogged()) { return; }
            if (PublishingHelpers.blueberryTaskPaneExists())
            {
                if (!PublishingHelpers.blueberryTaskPaneVisible())
                {
                    PublishingHelpers.showBlueberryTaskPane();
                    return;
                }
                MessageBox.Show("Please use the control panel on the right hand side.");
              return;
            }
            publishBlueberryTaskPane = new BlueberryTaskPane();
            string currentWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            myTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(publishBlueberryTaskPane, "Publish" + currentWorkbook);
            myTaskPane.VisibleChanged += new EventHandler(myTaskPane_VisibleChanged);
            myTaskPane.Visible = true;
            
        }

        private void myTaskPane_VisibleChanged(object sender, System.EventArgs e)
        {

        }


        private void Update_Click(object sender, RibbonControlEventArgs e)
        {
            if (!UserManagement.userLogged()) { return; }
            Dictionary<string, dynamic> publishedData = Publishing.getPublished();
            if (publishedData.Count == 0) { MessageBox.Show("No data was published from this workbook therefore, there is nothing to be updated."); return; }
            if (!PublishingHelpers.validateUpdateRanges(publishedData)) { MessageBox.Show("One or some of the ranges to be updated are empty"); return; }
            publishedData = PublishingHelpers.selectValidUpdateWorksheets(publishedData);

            try
            {
                MessageBox.Show(PublishingHelpers.publishSeveral(publishedData));
            }
            catch (KeyNotFoundException ex)
            {
                return;
            }
        }

        private void Fetch_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton senderObject = (RibbonButton)sender;
            string senderLabel = senderObject.Label;
            if (!UserManagement.userLogged()) { return; }
            if (FetchingHelpers.validateIDpreFetch())
            {
                Dictionary<string, dynamic> fetchedData = Fetching.fetchData();
                if (fetchedData.Count == 0) { return; }
                if (FetchingHelpers.validateIDPostFetch(fetchedData, senderLabel))
                {
                    FetchingHelpers.saveToExcel(fetchedData);
                }
            }
        }

        private void Refresh_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton senderObject = (RibbonButton)sender;
            string senderLabel = senderObject.Label;
            string errorMessage = "";
            if (!UserManagement.userLogged()) { return; }
            try
            {
                Dictionary<string, dynamic> fetchedData = Fetching.getFetched();
                int fetchDataItemsCount = fetchedData["names"].Count;
                for (int i = 0; i < fetchDataItemsCount; i++)
                {
                    Dictionary<string, dynamic> singleResult = new Dictionary<string, dynamic>();
                    singleResult.Add("bapi_id", fetchedData["bapi_ids"][i]);
                    singleResult.Add("user", fetchedData["users"][i]);
                    singleResult.Add("description", fetchedData["descriptions"][i]);
                    singleResult.Add("organization", fetchedData["organizations"][i]);
                    singleResult.Add("workbook_path", fetchedData["workbook_paths"][i]);
                    singleResult.Add("workbook", fetchedData["workbooks"][i]);
                    singleResult.Add("worksheet", fetchedData["worksheets"][i]);
                    singleResult.Add("destination_cell", fetchedData["destination_cells"][i]);
                    Dictionary<string, dynamic> toBeSaved = Fetching.fetchData(singleResult);
                    if (FetchingHelpers.validateIDPostFetch(toBeSaved, senderLabel))
                    {
                        FetchingHelpers.saveToExcel(toBeSaved);
                    }
                    else
                    {
                        errorMessage = "One or more data points which you are trying to fetch doesn't exist anymore.";
                    }
                }
                if (errorMessage != "") { MessageBox.Show(errorMessage); }
            }
            catch (KeyNotFoundException ex)
            {
                MessageBox.Show("There was no data downloaded in this worksheet yet.");
                return;
            }
        }

        /* Garbage collections methods */
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void GoToWebPlatformButton_Click(object sender, RibbonControlEventArgs e)
        {
            System.Diagnostics.Process.Start("http://blueberry-api.appspot.com");
        }

        private void LogInButton_Click(object sender, RibbonControlEventArgs e)
        {
            string username = usernameBox.Text;
            string password = passwordBox.Text;
            string responseSessionCookie;
            responseSessionCookie = "";

            var request = (HttpWebRequest)WebRequest.Create(GlobalVariables.blueberryAPIurl + "login");

            var postData = "email=" + username;
            postData += "&password=" + password;
            var data = Encoding.ASCII.GetBytes(postData);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;

            object[] httpResponseArgs = new object[] { "HttpResponse" };
            BlueberryHTTPResponse httpResponse = new BlueberryHTTPResponse(request, data, httpResponseArgs);

            responseSessionCookie = httpResponse.sendHTTPRequest(new BlueberryHTTPResponse.handleResponseDelegate(LogInButton_ClickHandleResponse),
                new BlueberryHTTPResponse.handleReponseExceptionsDelegate(LogInButton_ClickHandleExceptions));
            
            // If the user was authenticated, the response will have a session ID inside "Set-Cookie" header.
            if (responseSessionCookie == null)
            {
                MessageBox.Show("Invalid username or password");
            }
            else
            {
                string sessionCookieValueTemp = responseSessionCookie.Split(';')[0];
                Dictionary<string, string> sessionCookie = new Dictionary<string, string>();
                sessionCookie.Add("auth", Regex.Split(sessionCookieValueTemp, "auth=")[1]);

                GlobalVariables.sessionData = sessionCookie;
                GlobalVariables.sessionData.Add("loggedUser", usernameBox.Text);

                LogInButton.Visible = false;
                usernameBox.Visible = false;
                passwordBox.Visible = false;
                LogOutButton.Visible = true;
            }
            
            usernameBox.Text = "";
            passwordBox.Text = "";
        }

        private static dynamic LogInButton_ClickHandleResponse(object[] args)
        {
            HttpWebResponse httpWebResponse = (HttpWebResponse)args[0];
            return httpWebResponse.Headers["Set-Cookie"];
        }

        private static dynamic LogInButton_ClickHandleExceptions(object[] args)
        {
            MessageBox.Show("Please connect to Internet.");
            return null;
        }



        private void LogOutButton_Click(object sender, RibbonControlEventArgs e)
        {
            string url = GlobalVariables.blueberryAPIurl + "logout";

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

            request.Method = "GET";
            request.Headers.Add("Cookie", "auth=" + GlobalVariables.sessionData["auth"]);

            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    GlobalVariables.sessionData.Remove("auth");
                    GlobalVariables.sessionData.Remove("loggedUser");
                    LogInButton.Visible = true;
                    usernameBox.Visible = true;
                    passwordBox.Visible = true;
                    LogOutButton.Visible = false;
                }
            }
            catch (WebException ex)
            {
                if ((ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                    || ex.Status == WebExceptionStatus.ConnectFailure
                    || ex.Status == WebExceptionStatus.NameResolutionFailure)
                {
                    MessageBox.Show("Please connect to Internet.");
                }
                else
                {
                    throw;
                }
            }
        }


        private void TestButton_Click(object sender, RibbonControlEventArgs e)
        {
            ProxyFactory factory = new ProxyFactory(new Utils.ServiceCommand());
            factory.AddAdvice(new Utils.ConsoleLoggingAroundAdvice());
            Utils.ICommand command = (Utils.ICommand)factory.GetProxy();
            command.Execute("This is the argument");
        }

    }
}
