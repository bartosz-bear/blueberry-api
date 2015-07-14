using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Web.Script.Serialization;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;
using PublishingHelpers = ExcelAddIn1.Controllers.Helpers.PublishingHelpers;
using Publishing = ExcelAddIn1.Controllers.Publishing;
using ExcelAddIn1.Utils;

namespace ExcelAddIn1
{
    public partial class BlueberryTaskPane : UserControl
    {
        public BlueberryTaskPane()
        {
            InitializeComponent();
        }

        private void PublishButton_Click(object sender, EventArgs e)
        {
            if (!UserManagement.userLogged()) { return; }
            string validationResult = PublishingHelpers.validatePublishingInputs();
            if (validationResult == "Pass")
            {
                Publishing.publishData();
            }
            else
            {
                MessageBox.Show(validationResult);
            }

        }
    }
}
