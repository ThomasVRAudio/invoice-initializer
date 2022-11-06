using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Invoice_Initializer
{
    public class ClientBook
    {
        public string? clientCompany;
        public string? clientAdress;
        public string? clientZipCode;
        public string? clientProvince;
        public string? clientEmail;
        public int? clientID;

        public void GetInfo(int id, string path)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            clientCompany = xlRange.Cells[2, id + 1].Value2.ToString();
            clientAdress = xlRange.Cells[3, id + 1].Value2.ToString();
            clientZipCode = xlRange.Cells[4, id + 1].Value2.ToString();
            clientProvince = xlRange.Cells[5, id + 1].Value2.ToString();
            clientEmail = xlRange.Cells[6, id + 1].Value2.ToString();
            clientID = id;
           

            _ = Marshal.ReleaseComObject(xlRange);
            _ = Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            _ = Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            _ = Marshal.ReleaseComObject(xlApp);


        }
    }
}
