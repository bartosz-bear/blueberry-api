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
using PublishingValidators = ExcelAddIn1.Controllers.Validators.PublishingValidators;
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
            Excel.Workbook xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet xlWorkSheet  = (Excel.Worksheet)xlWorkBook.ActiveSheet;
            Excel.Range xlRange = (Excel.Range)xlWorkSheet.Application.Selection;
            string xlName = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingNameTextBox.Text;
            string xlDescription = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingDescriptionTextBox.Text;
            string xlOrganization = Globals.Ribbons.Ribbon1.publishBlueberryTaskPane.PublishingOrganizationTextBox.Text;
            string xlDataOwner = GlobalVariables.sessionData["loggedUser"];
            PublishingValidators validator = new PublishingValidators(xlRange, xlName, xlDescription, xlOrganization, xlDataOwner);
            string validationResult = validator.validatePublishingInputs(new List<string> {"isPublishRangeEmpty",
                                                                         "isAnyBlueberryTaskPaneFieldEmpty",
                                                                         "areInputsSpecialCharactersFree",
                                                                         "isIDUsed"});
            //string validationResult = PublishingHelpers.validatePublishingInputs();
            if (validationResult == "Pass")
            {
                Publishing.publishData();
                MessageBox.Show("Data has been published.");
            }
            else
            {
                MessageBox.Show(validationResult);
            }

        }
    }
}
