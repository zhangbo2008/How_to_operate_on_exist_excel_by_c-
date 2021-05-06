using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;


using Excel = Microsoft.Office.Interop.Excel;
namespace Microsoft.Office.Interop

{
    class Program
    {
        static void Main(string[] args)
        {


            Excel._Application objExcel;

            objExcel = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");


         //   objExcel.Workbooks.Open("c:/1.xlsx");
            objExcel.Workbooks.Close();
            objExcel.Workbooks.Open("c://1.xlsx");



            System.Console.Write("dsafdsfsad");
        }
















    }
}
