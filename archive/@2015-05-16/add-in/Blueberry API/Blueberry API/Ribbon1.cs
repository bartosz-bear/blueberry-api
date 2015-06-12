using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Runtime.InteropServices;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;   

namespace Blueberry_API
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range xlRange;

            xlWorkBook = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.ActiveSheet;

            xlRange = xlWorkSheet.UsedRange;

            

            ArrayList publishingList = new ArrayList();

            //string str;
            int rCnt = 0;
            int cCnt = 0;

            for (rCnt = 1; rCnt <= xlRange.Rows.Count; rCnt++)
            {
                for (cCnt = 1; cCnt <= xlRange.Columns.Count; cCnt++)
                {
                    publishingList.Add((string)(xlRange.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    //str = (string)(range.Cells[rCnt, cCnt] as Excel.Range).Value2;
                    //MessageBox.Show(str);
                }
            }

            MessageBox.Show((string)publishingList[0]);
            MessageBox.Show((string)publishingList[1]);


            /*
            string str;
            str = xlRange.Value2;
            MessageBox.Show(str);
            string str2;
            str2 = xlRange.Value2;
            MessageBox.Show(str2);

            

            */




            /// This is a working code which creates a new file and saves data
            /*
            excel.application xlapp = new microsoft.office.interop.excel.application();

            if (xlapp == null)
            {
                messagebox.show("excel is not properly installed!!");
                return;
            }

            
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Sheet 1 content";

            xlWorkBook.SaveAs("C:\\Users\\chbapie\\Desktop\\Bartosz\\apiquitous\\spreadsheets\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file d:\\csharp-Excel.xls");
            */

        }

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
    }
}
