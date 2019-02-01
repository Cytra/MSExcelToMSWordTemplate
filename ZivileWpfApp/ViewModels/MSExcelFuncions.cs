using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows;
using DataTable = System.Data.DataTable;

namespace ZivileWpfApp.ViewModels
{
    public static class MSExcelFuncions
    {

        public static DataTable ReadExcelDOcument(string fileLocation)
        {
            var result = new DataTable();
            result.Clear();
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileLocation);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            //DataTable dt = new DataTable();
            //dt.Clear();
            //dt.Columns.Add("Name");
            //dt.Columns.Add("Marks");
            //DataRow _ravi = dt.NewRow();
            //_ravi["Name"] = "ravi";
            //_ravi["Marks"] = "500";
            //dt.Rows.Add(_ravi);

            //Add Column Names
            for (int j = 1; j <= colCount; j++)
            {
                result.Columns.Add(xlRange.Cells[1, j].Value2.ToString());
            }


            //DataRow workRow;

            //for (int i = 0; i <= 9; i++)
            //{
            //    workRow = workTable.NewRow();
            //    workRow[0] = i;
            //    workRow[1] = "CustName" + i.ToString();
            //    workTable.Rows.Add(workRow);
            //}

            //Add Excel Data
            for (int i = 2; i <= rowCount; i++)
            {
                DataRow row = result.NewRow();
                for (int j = 1; j <= colCount; j++)
                {
                    var value = xlRange.Cells[i, j].Value;
                    if (value != null)
                    {
                        row[j - 1] = xlRange.Cells[i, j].Value.ToString();
                    }
                    //MessageBox.Show(xlRange.Cells[i, j].Value2.ToString());
                    
                }
                result.Rows.Add(row);
            }

            return result;
        }
    }
}
