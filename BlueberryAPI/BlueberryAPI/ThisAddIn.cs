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
    public partial class ThisAddIn
    {
        //private BlueberryTaskPane taskPaneControl1;
        //private Microsoft.Office.Tools.CustomTaskPane taskPaneValue;


        public BlueberryTaskPane publishBlueberryTaskPane;
        public Microsoft.Office.Tools.CustomTaskPane myTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(
                Application_WorkbookActivate);

            MessageBox.Show("Happening");

            publishBlueberryTaskPane = new BlueberryTaskPane();
            myTaskPane = this.CustomTaskPanes.Add(
                publishBlueberryTaskPane, "Publish");
            publishBlueberryTaskPane.Visible = true;
            myTaskPane.VisibleChanged += new EventHandler(myTaskPane_VisibleChanged);
            //MessageBox.Show("You can see it");
            
            //taskPaneValue.VisibleChanged +=
            //    new EventHandler(taskPaneValue_VisibleChanged);
        }



        void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            myTaskPane.Visible = true;
            //if (Wb.Name == "Book1.xlsx")
            //    myTaskPane.Visible = true;
            //else
            //    myTaskPane.Visible = false;
        }

        private void myTaskPane_VisibleChanged(object sender, System.EventArgs e)
        {
            //Globals.Ribbons.Ribbon1.toggleButton1.Checked = myTaskPane.Visible;
            //myTaskPane.Visible = true;
            MessageBox.Show("Clicked 2");
        }

        public Microsoft.Office.Tools.CustomTaskPane MyTaskPane
        {
            get
            {
                return this.myTaskPane;
            }
            set
            {
                myTaskPane = value;
            }
        }

        /*

        private void taskPaneValue_VisibleChanged(object sender, System.EventArgs e)
        {
            Console.WriteLine("Working");
            //Globals.Ribbons.Ribbon1.button2.Checked =
            //    taskPaneValue.Visible;
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return taskPaneValue;
            }
        }

         */ 
        
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
