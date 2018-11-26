using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Accord.MachineLearning;
using Accord.Controls;

namespace logInterpret
{
    public partial class wellLogInterprete : Form
    {
        public wellLogInterprete()
        {
            InitializeComponent();
        }

        string fileName = "";
        string datDir = "";
        OpenFileDialog loadDat=null;
        DataSet ds = new DataSet();

        private void btnOpen_Click(object sender, EventArgs e)
        {
            loadDat = new OpenFileDialog();
            loadDat.Multiselect = true;
            loadDat.Filter = "xls files|*.xls|xlsx file|*.xlsx|txt file|*.txt|dat file|*.dat|All files(*.*)|*.*";
            loadDat.FilterIndex = 2;
            loadDat.RestoreDirectory = true;

            if (loadDat.ShowDialog() == DialogResult.OK)
            {
                fileName = System.IO.Path.GetFullPath(loadDat.FileName);
                datDir = System.IO.Path.GetDirectoryName(loadDat.FileName);
                txtFilePath.Text = fileName;
                //txtDatPath.Text = datDir;
                btnRun.Enabled = true;
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            ds = new DataSet();
            if (fileName.Contains(".xls") || fileName.Contains(".xlsx"))
                ds = excelTrackle.LoadDataFromExcel(fileName, "Sheet1");
            else if (fileName.Contains(".txt") || fileName.Contains(".dat"))
                ds.Tables.Add(excelTrackle.LoadDataFromText(fileName));
            //for (int ti = 0; ti < ds.Tables.Count; ti++)
            //    for (int i = 0; i < ds.Tables[ti].Rows.Count; i++)
            //    {
            //        for (int j = 0; j < ds.Tables[ti].Columns.Count; j++)
            //            Console.Write(ds.Tables[ti].Rows[i][j].ToString() + "\t");
            //        Console.Write("\n");
            //    }
            double[][] data = new double[ds.Tables[0].Rows.Count - 1][];
            for (int i = 0; i < data.Length; i++)
            {
                data[i] = new double[2] { Convert.ToDouble(ds.Tables[0].Rows[i + 1][2]), Convert.ToDouble(ds.Tables[0].Rows[i + 1][3]) };
            }

            KMeans kmean = new KMeans(2);
            int[] result = kmean.Compute(data);
            ds.Tables.Add();
            ds.Tables[1].Columns.Add("depth", Type.GetType("System.Double"));
            ds.Tables[1].Columns.Add("lith1", Type.GetType("System.Int32"));
            DataRow dr = ds.Tables[1].NewRow();
            for (int i = 0; i < result.Length; i++)
            {
                dr = ds.Tables[1].NewRow();
                dr["depth"] = ds.Tables[0].Rows[i + 1][0];
                dr["lith1"] = result[i];
                ds.Tables[1].Rows.Add(dr);
            }
            Console.WriteLine("log interpretion is OK");
            ScatterplotBox.Show("lith", data, result);
            btnSave.Enabled = true;

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "text(*.txt)|*.txt|dat(*.dat)|*.dat|2003excel(*.xls)|*.xls|2007excel(*.xlsx)|*.xlsx";
            saveFile.RestoreDirectory = true;

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                string fileName = saveFile.FileName.ToString();
                if (fileName.Contains(".xls") || fileName.Contains(".xlsx"))
                    excelTrackle.ExportExcel(ds.Tables[1], fileName);
                else if (fileName.Contains(".txt") || fileName.Contains(".dat"))
                    excelTrackle.ExportText(ds.Tables[1], fileName);
                    
            }
        }

        private void btnQuit_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private static double[][] arrayFromTable (DataTable table)
        {
            double[][] data = new double[table.Rows.Count][];
            for(int i=1;i<table.Rows.Count;i++)
            {
                data[i] = new double[2]{ Convert.ToDouble(table.Rows[i][1]), Convert.ToDouble(table.Rows[i][2])};
                
            }
            return data;
        }

        private void txtFilePath_TextChanged(object sender, EventArgs e)
        {
            fileName = txtFilePath.Text;
        }



    }

    

}
