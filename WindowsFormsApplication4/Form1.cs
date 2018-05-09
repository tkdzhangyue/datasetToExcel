using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office;
using System.IO;
using System.Reflection;


namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable("Table");

            dt.Columns.Add(new DataColumn("id", typeof(int)));
            dt.Columns.Add(new DataColumn("name", typeof(string)));
            dt.Columns.Add(new DataColumn("img", typeof(string)));
            dt.Columns.Add(new DataColumn("timer", typeof(string)));

            DataRow dr = dt.NewRow();
            dr["id"] = 1;
            dr["name"] = "AA";
            dr["timer"] ="now";
            dr["img"] = "~/img/1.png";
            dt.Rows.Add(dr);

            ds.Tables.Add(dt);


            this.dataGridView1.DataSource = ds.Tables[0];

            DataSetToExcel(ds, false);


        }





        /// <summary>
        /// 将数据集中的数据导出到EXCEL文件
        /// </summary>
        /// <param name="dataSet">输入数据集</param>
        /// <param name="isShowExcle">是否显示该EXCEL文件</param>
        /// <returns></returns>
        public bool DataSetToExcel(DataSet dataSet, bool isShowExcle)
        {
            DataTable dataTable = dataSet.Tables[0];
            int rowNumber = dataTable.Rows.Count;//不包括字段名
            int columnNumber = dataTable.Columns.Count;
            int colIndex = 0;

            if (rowNumber == 0)
            {
                return false;
            }

            //建立Excel对象 
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            //excel.Application.Workbooks.Add(true);
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];
            excel.Visible = isShowExcle;
            Microsoft.Office.Interop.Excel.Range range;

            //生成字段名称 
            foreach (DataColumn col in dataTable.Columns)
            {
                colIndex++;
                excel.Cells[1, colIndex] = col.ColumnName;
            }

            object[,] objData = new object[rowNumber, columnNumber];

            for (int r = 0; r < rowNumber; r++)
            {
                for (int c = 0; c < columnNumber; c++)
                {
                    objData[r, c] = dataTable.Rows[r][c];
                }
                //Application.DoEvents();
            }

            // 写入Excel 
            //range = worksheet.get_Range(excel.Cells[2, 1], excel.Cells[rowNumber + 1, columnNumber]);
            Microsoft.Office.Interop.Excel.Range c1 = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[2, 1];
            Microsoft.Office.Interop.Excel.Range c2 = (Microsoft.Office.Interop.Excel.Range)worksheet.Cells[rowNumber + 1, columnNumber];
            range = worksheet.get_Range(c1, c2);
            range.Value2 = objData;

            //屏蔽掉系统跳出的Alert
            excel.AlertBeforeOverwriting = false;

            //获取你使用的excel 的版本号
            string Version = excel.Version;
            Double FormatNum;
            //使用Excel 97-2003
            if (Convert.ToDouble(Version) < 12)
            {
                FormatNum = -4143;
            }
            //使用excel 2007或者更新de 
            else
            {
                FormatNum = 56;
            }

            //初始化文件名
            DateTime excelNow = DateTime.Now;
            String filePath = Convert.ToString(excelNow.Year) + "_" + Convert.ToString(excelNow.Month) + "_" + Convert.ToString(excelNow.Day) + "_" + Convert.ToString(excelNow.Hour) + "_" + Convert.ToString(excelNow.Minute) + "_" + Convert.ToString(excelNow.Second) + ".xls";

            //保存到指定目录
            //workbook.SaveAs(filePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            workbook.SaveAs(filePath, FormatNum);

            excel.Quit();
            //释放掉多余的excel进程
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            excel = null;

            return true;
        }
        }

}
