﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace insta_json_to_excel
{
    public partial class Form1 : Form
    {
        DataTable dsexcel = new DataTable();
        DataTable dtall = new DataTable();
        public Form1()
        {
            InitializeComponent();
        }


        private void button_Click(object sender, EventArgs e)
        {
            dtall.Columns.Add(new DataColumn("href", typeof(string)));
            dtall.Columns.Add(new DataColumn("type", typeof(string)));
            dtall.Columns.Add(new DataColumn("time", typeof(string)));


            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filePath = file.FileName;
                string filename = file.SafeFileName;

                filePath = filePath.Replace(filename, "");

                string instauser = filePath.Replace("\\followers_and_following\\", "");
                int l = instauser.LastIndexOf("\\");
                instauser = instauser.Substring(l + 1);

                DirectoryInfo df = new DirectoryInfo(@filePath);
                FileInfo[] Files = df.GetFiles("*.json"); //Getting Text files
                string str = "";

                Process[] process = Process.GetProcessesByName("Excel");
                foreach (System.Diagnostics.Process p in process)
                {
                    if (p.MainWindowTitle.Length == 0)
                    {
                        p.Kill();
                    }
                }

                string w = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + instauser + ".xlsx";
                var excel = new Excel.Application();
                var wrbk = excel.Workbooks.Add(Type.Missing);
                wrbk.SaveAs(w, 51, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlExclusive, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
                wrbk = excel.Workbooks.Open(@w);



                foreach (FileInfo file1 in Files)
                {

                    string f1 = File.ReadAllText(filePath + file1.Name)

                 .Replace("\"relationships_followers\":", "")
                 .Replace("\"relationships_following\":", "")
                 .Replace("\"relationships_following_hashtags\":", "")
                 .Replace("\"relationships_follow_requests_sent\":", "")
                 .Replace("\"relationships_permanent_follow_requests\":", "")
                 .Replace("\"relationships_unfollowed_users\":", "")
                 .Replace("\"relationships_dismissed_suggested_users\":", "")
                 .Replace("\"media_list_data\": [", "")
                 .Replace("\n", "")
                 .Replace("\"media_list_data\": [              ],", "")
                 .Replace("\"title\": \"\",", "")
                 .Replace("\"string_list_data\": [        {          ", "")
                 .Replace("    \"string_list_data\": [      {       ", "")
                 .Replace("[  {                  ", "")
                 .Replace("},  {", "},    {")
                 .Replace("{   [    {                  ", "")
                 .Replace("],", "")
                 .Replace("]", "");

                    int idqx = f1.IndexOf(",");
                    string f2 = f1.Substring(0, idqx + 1);

                    DataTable dt2 = new DataTable();

                    string[] jsonStringArray = Regex.Split(f1, "},    {");
                    List<string> ColumnsName = new List<string>();
                    foreach (string jSA in jsonStringArray)
                    {
                        string[] jsonStringData = Regex.Split(jSA, ",");
                        foreach (string ColumnsNameData in jsonStringData)
                        {

                            try
                            {
                                int idx = ColumnsNameData.IndexOf(":");
                                string ColumnsNameString = ColumnsNameData.Substring(0, idx - 1).Replace("\"", "").Trim();
                                ColumnsName.Add(ColumnsNameString);

                            }
                            catch (Exception ex)
                            {

                            }

                        }
                        break;

                    }

                    foreach (string AddColumnName in ColumnsName)
                    {
                        dt2.Columns.Add(AddColumnName.Trim());

                    }
                    foreach (string jSA in jsonStringArray)
                    {
                        if (jSA != " ")
                        {
                            string[] RowData = Regex.Split(jSA.Replace("{", "").Replace("}", ""), ",");
                            DataRow nr = dt2.NewRow();

                            foreach (string rowData in RowData)
                            {
                                try
                                {
                                    int idx = rowData.IndexOf(":");
                                    string RowColumns = rowData.Substring(0, idx - 1).Replace("\"", "").Trim();
                                    string RowDataString = rowData.Substring(idx + 1).Replace("\\n", "").Replace("\"", "").Replace("\\", "").Trim();
                                    if (RowColumns == "timestamp")
                                    {
                                        DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);
                                        RowDataString = dateTime.AddSeconds(Convert.ToDouble(RowDataString)).ToLocalTime().ToString("MM/dd/yyyy HH:mm:ss");
                                    }
                                    nr[RowColumns] = RowDataString;
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                            dt2.Rows.Add(nr);
                        }
                    }

                    AddExcelSheet(dt2, wrbk, file1.Name);

                }

                dtall.DefaultView.Sort = "href";
                dtall = dtall.DefaultView.ToTable();
                DataRow[] r = dtall.Select("type ='followers_1'");

                foreach (DataRow row in r)
                {
                    DataRow[] dtr = dtall.Select("href ='" + row[0].ToString() + "'"); //name is the column in the data table
                    foreach (var drow in dtr)
                    {
                        drow.Delete();
                    }
                    dtall.AcceptChanges();
                }


                DataRow[] r1 = dtall.Select("type ='removed_suggestions'");

                foreach (DataRow row in r1)
                {
                    DataRow[] dtr = dtall.Select("href ='" + row[0].ToString() + "'"); //name is the column in the data table
                    foreach (var drow in dtr)
                    {
                        drow.Delete();
                    }
                    dtall.AcceptChanges();
                }

                DataRow[] r2 = dtall.Select("type ='following_hashtags'");

                foreach (DataRow row in r2)
                {
                    DataRow[] dtr = dtall.Select("href ='" + row[0].ToString() + "'"); //name is the column in the data table
                    foreach (var drow in dtr)
                    {
                        drow.Delete();
                    }
                    dtall.AcceptChanges();
                }

                AddExcelSheet(dtall, wrbk, "all");
                wrbk.Save();
                wrbk.Close();
                Process.Start(w);
            }
        }
        private void AddExcelSheet(DataTable dt, Excel.Workbook wb, string name)
        {
            if (dt.Rows.Count > 0)
            {
                Excel.Sheets sh = wb.Sheets;
                Excel.Worksheet osheet = sh.Add();
                osheet.Name = name.Replace(".json", "");
                int colIndex = 0;
                int rowIndex = 1;

                foreach (DataColumn dc in dt.Columns)
                {
                    colIndex++;
                    osheet.Cells[1, colIndex] = dc.ColumnName;
                }
                foreach (DataRow dr in dt.Rows)
                {
                    rowIndex++;
                    colIndex = 0;
                    DataRow nr = dtall.NewRow();
                    foreach (DataColumn dc in dt.Columns)
                    {
                        colIndex++;
                        osheet.Cells[rowIndex, colIndex] = dr[dc.ColumnName];

                        if (name == "all")
                        {

                        }
                        else if (dc.ColumnName == "href")
                        {

                            nr["href"] = dr[dc.ColumnName];
                            nr["type"] = name.Replace(".json", "");

                        }
                        else if (dc.ColumnName == "timestamp")
                        {
                            nr["time"] = dr[dc.ColumnName];
                            dtall.Rows.Add(nr);
                        }

                    }
                }
                osheet.Columns.AutoFit();
                osheet.Rows.AutoFit();
            }
        }
    }
}
