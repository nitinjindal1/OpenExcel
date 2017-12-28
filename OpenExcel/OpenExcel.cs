using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;
using System.Text.RegularExpressions;
using System.Reflection;
using Microsoft.Office.Core; //Added to Project Settings' References from C:\Program Files (x86)\Microsoft Visual Studio 10.0\Visual Studio Tools for Office\PIA\Office14 - "office"

namespace OpenExcel
{
    public partial class OpenExcel : Form
    {
        //Create a Workbook Object
        public static Excel.Workbook MyBook = null;

        //Create an Application Object
        public static Excel.Application MyApp = null;

        //Create a Worksheet Object
        public static Excel.Worksheet MySheet = null;

        //String variable for Project Id
        String Project_Id = null;

        object misValue = System.Reflection.Missing.Value;

        public OpenExcel()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\ub765xj\Downloads\PartnerData.xlsm", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            //Save Project Id in a variable
            Project_Id = xlWorkSheet.get_Range("B2", "B2").Value2.ToString();
          

            //Message Box showing Project Id
            MessageBox.Show(xlWorkSheet.get_Range("B2", "B2").Value2.ToString());

            xlWorkBook.SaveAs(@"C:\Users\ub765xj\Load Test\Export\"+Project_Id, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
                                    
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            object oMiss = System.Reflection.Missing.Value;

            xlApp = new Excel.Application();
            xlApp.DisplayAlerts = true;
            xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityByUI;
            // Make the excel visible
            xlApp.Visible = true;
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\ub765xj\Documents\CJV 2.0\Performance Test\CJV Bulk Data Import\Export\PartnerData_27-Dec-17.xlsm", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);



            //Enable Macros
            if (xlWorkBook.HasVBProject)  // Has macros
            {
                try
                {
                    // Show "Microsoft Excel Security Notice" prompt
                    var project = xlWorkBook.VBProject;
                }
                catch (System.Runtime.InteropServices.COMException comex)
                {
                    // Macro is enabled.
                }
            }

            xlWorkBook.SaveAs(@"C:\Users\ub765xj\Documents\CJV 2.0\Performance Test\CJV Bulk Data Import\Import\" + xlWorkBook.Name, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            MessageBox.Show("Macros Enabled");
            xlApp.Quit();


            xlWorkBook.Close(true, misValue, misValue);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            xlWorkBook = null;
            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
            xlApp = null;

            //Add Rows for Partner Id's


            //Save Project Id in a variable
            // Project_Id = xlWorkSheet.get_Range("B2", "B2").Value2.ToString();


            //Message Box showing Project Id
            //MessageBox.Show(xlWorkSheet.get_Range("B2", "B2").Value2.ToString());

            // xlWorkBook.SaveAs(@"C:\Users\ub765xj\Documents\CJV 2.0\Performance Test\CJV Bulk Data Import\Import\" + xlWorkBook.Name, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            // xlWorkBook.Close(true, misValue, misValue);
            // MessageBox.Show("Macros Enabled");
            // xlApp.Quit();

            // releaseObject(xlWorkSheet);
            // releaseObject(xlWorkBook);
            // releaseObject(xlApp);

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

        private void button3_Click(object sender, EventArgs e)
        {


            Excel.Worksheet oWorksheet;
            object misValue = System.Reflection.Missing.Value;
                object oMiss = System.Reflection.Missing.Value;

                // Object for missing (or optional) arguments.
                object oMissing = System.Reflection.Missing.Value;


                // Create an instance of Microsoft Excel, make it visible,
                // and open Book1.xls.
                Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
                oExcel.Visible = true;
                Microsoft.Office.Interop.Excel.Workbooks oBooks = oExcel.Workbooks;
                            
            Microsoft.Office.Interop.Excel._Workbook oBook = null;
     
            oExcel.DisplayAlerts = true;
                oExcel.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityByUI;
                // Make the excel visible
                oExcel.Visible = true;
            
            // Get the file path
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = path + "\\PartnerData_27-Dec-17.xlsm";

            oBook = oExcel.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                oWorksheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);
            // Run the macro, "AddPartner"
            RunMacro(oExcel, new Object[] { "PartnerData_27-Dec-17.xlsm!AddPartner2" });
                                       
            // Quit Excel and clean up.
            oBook.Close(false, oMissing, oMissing);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
            oBook = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
            oBooks = null;
            oExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
            oExcel = null;

            //Garbage collection
            GC.Collect();

           
        }
          
        private void button4_Click(object sender, EventArgs e)
        {
            // Object for missing (or optional) arguments.
            object oMissing = System.Reflection.Missing.Value;

            // Create an instance of Microsoft Excel
            Excel.Application oExcel = new Excel.Application();

            // Make it visible
            oExcel.Visible = true;

            // Define Workbooks
            Excel.Workbooks oBooks = oExcel.Workbooks;
            Excel._Workbook oBook = null;
            Excel.Worksheet oSheet;
            oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(2);

            // Get the file path
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            path = path + "\\PartnerData_27-Dec-17.xlsm";

            //Open the file, using the 'path' variable
            oBook = oBooks.Open(path, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);


            oExcel.Run("!AddPartner2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);



            // Run the macro, "First_Macro"
            //RunMacro(oBook, new Object[] { "PartnerData_27-Dec-17.xlsm!AddPartner2" });

            // Quit Excel and clean up.
            oBook.Close(false, oMissing, oMissing);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
            oBook = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
            oBooks = null;
            oExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
            oExcel = null;

            //Garbage collection
            GC.Collect();
        }

        private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default | System.Reflection.BindingFlags.InvokeMethod, null, oApp, oRunArgs);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook = null;
            xlApp.DisplayAlerts = true;
            Excel.Range range;
            xlApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityByUI;
            // Make the excel visible
            xlApp.Visible = true;
           

            try
            {
                //Start Excel and open the workbook.
                xlWorkBook = xlApp.Workbooks.Open(@"C:\Users\ub765xj\Documents\CJV 2.0\Performance Test\CJV Bulk Data Import\Export\PartnerData_27-Dec-17.xlsm", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);


                //Run the macros by supplying the necessary arguments
                xlApp.Run("!AddPartner2", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                //xlWorkBook.Save();
                //Clean-up: Close the workbook
                //xlWorkBook.Close(false);

                //Quit the Excel Application
                //xlApp.Quit();
            }
            catch (Exception ex)
            {
            }
            finally
            {
                //~~> Clean Up
                releaseObject(xlApp);
                releaseObject(xlWorkBook);
            }
        }
    }
    }
