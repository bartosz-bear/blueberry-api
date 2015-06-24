using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelAddIn1
{

    public class GlobalVariables
    {
        public static Dictionary<string, string> sessionID;
        public static string blueberryAPIurl = "http://localhost.:8080/";
        //public static string blueberryAPIurl = "http://blueberry-api.appspot.com/";
    }

    public partial class ThisAddIn
    {


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }
      
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
