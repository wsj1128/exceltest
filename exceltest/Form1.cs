using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace exceltest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            Microsoft.Office.Interop.Excel.Application data = new Microsoft.Office.Interop.Excel.Application();
            string path = AppDomain.CurrentDomain.BaseDirectory;

            //data.Visible = true;
            Workbook book = data.Workbooks.Open(path + "data.xlsx");
            Worksheet sheet = book.Worksheets[1];
            Range range = sheet.get_Range("A1", "B4");

            range[1][1] = "Name";
            range[2][1] = "Marks"; 
            //data.Workbooks.
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Marks");
            DataRow row1 = dt.NewRow();
            row1["Name"] = "Tom";
            row1["Marks"] = "96";
            dt.Rows.Add(row1);
            DataRow row2 = dt.NewRow();
            row2["Name"] = "Jerry";
            row2["Marks"] = "91";
            dt.Rows.Add(row2);
            DataRow row3 = dt.NewRow();
            row3["Name"] = "Pooly";
            row3["Marks"] = "100";
            dt.Rows.Add(row3);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    range[j + 1][i + 2] = dt.Rows[i][j];//range对象利用[]操作符来取每个单元格，但是range[][]的顺序是range[列号][行号]
                }
            }

            book.Save();
            //book.SaveAs(path + "Test.xlsx");
            //Workbook book = data.Workbooks.Open(path + "Test.xlsx ");
            data.Quit();
        }
    }
}
