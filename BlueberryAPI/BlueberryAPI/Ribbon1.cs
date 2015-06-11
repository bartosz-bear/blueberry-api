using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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

namespace ExcelAddIn1
{
    public partial class Ribbon1
    {


        public static string blueberryAPIurl = "http://localhost.:8080/";
        //public static string blueberryAPIurl = "http://blueberry-api.appspot.com/";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            MessageBox.Show("Bartosz Piechnik");
        }

        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            MessageBox.Show("Bartosz");
        }

        /* Buttons Events */

        private void Publish_Click(object sender, RibbonControlEventArgs e)
        {
            //BlueberryTaskPane publishBlueberryTaskPane;
            //Microsoft.Office.Tools.CustomTaskPane myTaskPane;

            //Excel.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(
            //    Application_WorkbookActivate);

            //MessageBox.Show("Clicked 2");

            //publishBlueberryTaskPane = new BlueberryTaskPane();
            //myTaskPane = this.CustomTaskPanes.Add(
            //    publishBlueberryTaskPane, "Publish");

            Globals.ThisAddIn.MyTaskPane.Visible = true;

        }

        private void Update_Click(object sender, RibbonControlEventArgs e)
        {
            Dictionary<string, dynamic> publishedData = Publishing.getPublished();
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
        }

        private void Fetch_Click(object sender, RibbonControlEventArgs e)
        {
            FetchingHelpers.saveToExcel(Fetching.fetchData());
        }

        private void Refresh_Click(object sender, RibbonControlEventArgs e)
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
                FetchingHelpers.saveToExcel(Fetching.fetchData(singleResult));
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

        private void TestButton_Click(object sender, RibbonControlEventArgs e)
        {
            MessageBox.Show("Testing");
        }

    }
}
