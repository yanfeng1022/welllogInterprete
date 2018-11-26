using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Windows.Forms;

namespace logInterpret
{
    class excelTrackle
    {
        /// <summary>  
        /// 解析Excel，根据OleDbConnection直接连Excel  
        /// </summary>  
        /// <param name="fileName"></param>  
        /// <param name="name"></param>  
        /// <returns></returns>  
        public static DataSet LoadDataFromExcel(string fileName, string name)
        {
            try
            {
                string strConn;
                //  strConn = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source = " + filePath + ";Extended Properties=Excel 8.0";  
                strConn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0 Xml;HDR=No\"";  //这是2010的链接字符串，不同版本链接不同
                OleDbConnection OleConn = new OleDbConnection(strConn);
                OleConn.Open();
                string sql = "SELECT * FROM [" + name + "$]";//可是更改Sheet名称，比如sheet2，等等   
                OleDbDataAdapter OleDaExcel = new OleDbDataAdapter(sql, OleConn);
                DataSet OleDsExcle = new DataSet();
                OleDaExcel.Fill(OleDsExcle, name);
                OleConn.Close();
                return OleDsExcle;
            }
            catch (Exception err)
            {
                MessageBox.Show("数据绑定Excel失败! 失败原因：" + err.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }

        public static System.Data.DataTable LoadDataFromText(string fileName)
        {
            string[] ss = File.ReadAllLines(fileName);
            string[] dataName = ss[0].Split(new char[] { '\t', ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int i = 0; i < dataName.Length; i++)
                dt.Columns.Add(dataName[i], Type.GetType("System.Double"));
            for (int i = 1; i < ss.Length; i++)
            {
                string[] ssLine = ss[i].Split(new char[] { '\t', ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
                DataRow dr = dt.NewRow();
                for (int j = 0; j < ssLine.Length; j++)
                    dr[j] = Convert.ToDouble(ssLine[j]);
                dt.Rows.Add(ssLine);
            }
            return dt;
        }
 
        public static void ExportText(System.Data.DataTable table, string strTextFileName)
        {
            StreamWriter sw = new StreamWriter(strTextFileName);
            for (int j = 0; j < table.Columns.Count; j++)
                sw.Write(table.Columns[j].Caption + "\t");
            sw.Write("\n");
            for (int i = 1; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                    sw.Write(Convert.ToString(table.Rows[i][j]) + "\t");
                sw.Write("\n");
            }
            sw.Close();
        }
 

        #region
        /**/
        //// <summary>
        /// 导出 Excel 文件
        /// </summary>
        /// <param name="ds">要导出的DataSet</param>
        /// <param name="strExcelFileName">要导出的文件名</param>
        public static void ExportExcel(System.Data.DataTable table, string strExcelFileName)
        {
            object objOpt = Missing.Value;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            _Workbook wkb = excel.Workbooks.Add(objOpt);
            _Worksheet wks = (_Worksheet)wkb.ActiveSheet;

            wks.Visible = XlSheetVisibility.xlSheetVisible;

            int rowIndex = 1;
            int colIndex = 0;

            foreach (DataColumn col in table.Columns)
            {
                colIndex++;
                excel.Cells[1, colIndex] = col.ColumnName;
            }

            foreach (DataRow row in table.Rows)
            {
                rowIndex++;
                colIndex = 0;
                foreach (DataColumn col in table.Columns)
                {
                    colIndex++;
                    excel.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
                }
            }
            //excel.Sheets[0] = "sss";

            wkb.SaveAs(strExcelFileName, objOpt, null, null, false, false, XlSaveAsAccessMode.xlNoChange, null, null, null, null, null);
            wkb.Close(false, objOpt, objOpt);
            excel.Quit();
        }
        #endregion
    }
}
