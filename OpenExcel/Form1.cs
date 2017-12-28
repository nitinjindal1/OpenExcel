using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Core;

namespace OpenExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }




        private void RunMacro(object oApp, object[] oRunArgs)
        {
            oApp.GetType().InvokeMember("Run",
                System.Reflection.BindingFlags.Default |
                System.Reflection.BindingFlags.InvokeMethod,
                null, oApp, oRunArgs);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Object for missing (or optional) arguments.
            object oMissing = System.Reflection.Missing.Value;


            // Create an instance of Microsoft Excel, make it visible,
            // and open Book1.xls.
            Microsoft.Office.Interop.Excel.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            oExcel.Visible = true;
            Microsoft.Office.Interop.Excel.Workbooks oBooks = oExcel.Workbooks;
            Microsoft.Office.Interop.Excel._Workbook oBook = null;
            oBook = oBooks.Open("C:\\Users\\ub765xj\\Documents\\CJV 2.0\\Performance Test\\CJV Bulk Data Import\\Export\\PartnerData_27-Dec-17.xlsm", oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                    oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            // Run the macros.
            RunMacro(oExcel, new Object[] { "AddPartner2" });

            // Quit Excel and clean up.
            /*oBook.Close(false, oMissing, oMissing);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
            oBook = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
            oBooks = null;
            oExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
            oExcel = null;*/


        }

    }
}