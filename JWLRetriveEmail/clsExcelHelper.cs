using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data.SqlClient;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace JWLRetriveEmail
{
    class clsExcelHelper
    {
        public static DataSet ImportExcelXLSX(string Filepath, bool hasHeaders, bool FutureMapping = false)
        {
            clsCommon objCommon = new clsCommon();
            DataSet output = new DataSet();
            try
            {

                string HDR = (hasHeaders ? "Yes" : "No");
                // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=1\"";

                // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=1\"";

                //public static string path = @"C:\src\RedirectApplication\RedirectApplication\301s.xlsx";
                //  string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Filepath + ";Extended Properties=Excel 12.0;HDR=" + HDR + ";IMEX=1";
                string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=1\"";

                //string sql = "SELECT * FROM [Template$]";
                //string excelConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"";

                //using (OleDbDataAdapter adaptor = new OleDbDataAdapter(sql, excelConnection))
                //{
                //    DataSet ds = new DataSet();
                //    adaptor.Fill(ds);
                //}

                using (OleDbConnection conn = new OleDbConnection(strConn))
                {
                    conn.Open();

                    System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {
                null,
                null,
                null,
                "TABLE"
            });
                    int i = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        string sheet = string.Empty;
                        sheet = row["TABLE_NAME"].ToString();
                        OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                        cmd.CommandType = CommandType.Text;


                        System.Data.DataTable outputTable = new System.Data.DataTable(sheet);

                        output.Tables.Add(outputTable);

                        OleDbDataAdapter d = new OleDbDataAdapter(cmd);
                        try
                        {

                            d.Fill(outputTable);
                            if (i == 0)
                            {
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            // objCommon.LogEvents(ex.Message, "ImportExcelXLSX", System.Diagnostics.EventLogEntryType.Error, 1);
                            objCommon.WriteErrorLog(ex, "ImportExcelXLSX");
                        }
                        i++;
                    }
                }


                if (HDR == "No")
                {
                    foreach (DataColumn column in output.Tables[0].Columns)
                    {
                        string cName = output.Tables[0].Rows[0][column.ColumnName].ToString();
                        if (!output.Tables[0].Columns.Contains(cName) && cName != "")
                        {
                            column.ColumnName = cName;
                        }
                    }
                    output.Tables[0].Rows[0].Delete();
                    output.Tables[0].AcceptChanges();
                }

            }
            catch (Exception ex)
            {
                objCommon.WriteErrorLog(ex, "ImportExcelXLSX");
            }

            return output;
        }

        public static DataSet ImportExcelXLS(string Filepath, bool hasHeaders)
        {
            clsCommon objCommon = new clsCommon();
            string HDR = (hasHeaders ? "Yes" : "No");
            //  string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=1\"";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=1\"";


            DataSet output = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                conn.Open();

                System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {
                null,
                null,
                null,
                "TABLE"
            });

                foreach (DataRow row in dt.Rows)
                {
                    string sheet = row["TABLE_NAME"].ToString();

                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheet + "]", conn);
                    cmd.CommandType = CommandType.Text;


                    System.Data.DataTable outputTable = new System.Data.DataTable(sheet);

                    output.Tables.Add(outputTable);

                    OleDbDataAdapter d = new OleDbDataAdapter(cmd);
                    try
                    {
                        d.Fill(outputTable);

                    }
                    catch (Exception ex)
                    {
                        objCommon.WriteErrorLog(ex, "ImportExcelXLS");
                    }
                }
            }
            if (HDR == "No")
            {
                foreach (DataColumn column in output.Tables[0].Columns)
                {
                    string cName = output.Tables[0].Rows[0][column.ColumnName].ToString();
                    if (!output.Tables[0].Columns.Contains(cName) && cName != "")
                    {
                        column.ColumnName = cName;
                    }
                }
                output.Tables[0].Rows[0].Delete();
                output.Tables[0].AcceptChanges();
            }
            return output;
        }

        public static DataSet ImportExcelXLSXToDataSet(string Filepath, bool hasHeaders)
        {
            clsCommon objCommon = new clsCommon();
            string HDR = (hasHeaders ? "Yes" : "No");
            //  string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 8.0;HDR=" + HDR + ";IMEX=1\"";
            string strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Filepath + ";Extended Properties=\"Excel 12.0;HDR=" + HDR + ";IMEX=1\"";

            DataSet output = new DataSet();

            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                conn.Open();

                System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] {
                null,
                null,
                null,
                "TABLE"
            });
                // string sheet_name = "";

                // string sheet = "Sheet1$";

                foreach (DataRow row in dt.Rows)
                {
                    string sheet = row["TABLE_NAME"].ToString();

                    string sqlSelect = "SELECT * FROM [" + sheet + "]";
                    OleDbCommand cmd = new OleDbCommand(sqlSelect, conn);
                    cmd.CommandType = CommandType.Text;
                    System.Data.DataTable outputTable = new System.Data.DataTable(sheet);
                    output.Tables.Add(outputTable);

                    OleDbDataAdapter d = new OleDbDataAdapter(cmd);
                    try
                    {
                        d.Fill(outputTable);
                    }
                    catch (Exception ex)
                    {
                        objCommon.WriteErrorLog(ex, "ImportExcelXLSXToDataSet");
                    }
                }
            }
            if (HDR == "No")
            {
                if (output != null)
                {
                    if (output.Tables.Count > 0)
                    {
                        foreach (DataColumn column in output.Tables[0].Columns)
                        {
                            string cName = output.Tables[0].Rows[0][column.ColumnName].ToString();
                            if (!output.Tables[0].Columns.Contains(cName) && cName != "")
                            {
                                column.ColumnName = cName;
                            }
                        }

                        output.Tables[0].Rows[0].Delete();
                        output.Tables[0].AcceptChanges();
                    }
                    if (output.Tables.Count > 1)
                    {
                        foreach (DataColumn column in output.Tables[1].Columns)
                        {
                            string cName = output.Tables[1].Rows[0][column.ColumnName].ToString();
                            if (!output.Tables[1].Columns.Contains(cName) && cName != "")
                            {
                                column.ColumnName = cName;
                            }
                        }

                        output.Tables[1].Rows[0].Delete();
                        output.Tables[1].AcceptChanges();
                    }
                }
            }
            return output;
        }

        public static DataSet ImportCSV(string FileName, bool FirstRowHeader, string Delimiter, int SkipRows, System.Data.DataTable ColumnStruct = null)
        {
            clsCommon objCommon = new clsCommon();
            DataSet dsRet = new DataSet();
            System.Data.DataTable csvData = new System.Data.DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(FileName))
                {
                    csvReader.SetDelimiters(new string[] { Delimiter });
                    csvReader.HasFieldsEnclosedInQuotes = true;
                    for (int i = 0; i < SkipRows - 1; i++)
                    {
                        string[] fieldData = csvReader.ReadFields();
                    }

                    if (FirstRowHeader)
                    {
                        string[] colFields = csvReader.ReadFields();
                        foreach (string column in colFields)
                        {
                            DataColumn datecolumn = new DataColumn(column);
                            datecolumn.AllowDBNull = true;
                            csvData.Columns.Add(datecolumn);
                        }
                    }
                    else
                    {
                        foreach (DataColumn dc in ColumnStruct.Columns)
                        {
                            DataColumn datecolumn = new DataColumn(dc.ColumnName);
                            datecolumn.AllowDBNull = true;
                            csvData.Columns.Add(datecolumn);
                        }
                    }


                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        ////Making empty value as null
                        //for (int i = 0; i < fieldData.Length; i++)
                        //{
                        //    if (fieldData[i] == "")
                        //    {
                        //        fieldData[i] = null;
                        //    }
                        //}
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                objCommon.WriteErrorLog(ex, "ImportCSV");
            }
            dsRet.Tables.Add(csvData);
            return dsRet;
        }

        public static DataSet ImportCSVNew(string filePath)
        {
            clsCommon objCommon = new clsCommon();
            //reading all the lines(rows) from the file.
            string[] rows = File.ReadAllLines(filePath);
            DataSet dsRet = new DataSet();
            System.Data.DataTable dtData = new System.Data.DataTable();
            try
            {
                string[] rowValues = null;
                DataRow dr = dtData.NewRow();
                //Creating columns
                if (rows.Length > 0)
                {
                    int count = 1;
                    foreach (string columnName in rows[0].Split(','))
                    {
                        if (!dr.Table.Columns.Contains(columnName))
                        {
                            dtData.Columns.Add(columnName);
                        }
                        else
                        {
                            dtData.Columns.Add(columnName + "_" + count);
                            count++;
                        }
                    }
                }

                //Creating row for each line.(except the first line, which contain column names)
                for (int row = 1; row < rows.Length; row++)
                {
                    rowValues = rows[row].Split(',');
                    dr = dtData.NewRow();
                    dr.ItemArray = rowValues;
                    dtData.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                objCommon.WriteErrorLog(ex, "ImportCSVNew");
            }
            dsRet.Tables.Add(dtData);
            return dsRet;
        }

        public static DataSet ConvertCSVtoDataSet(string strFilePath)
        {
            clsCommon objCommon = new clsCommon();
            //reading all the lines(rows) from the file.
            DataSet dsRet = new DataSet();
            System.Data.DataTable dt = new System.Data.DataTable();
            try
            {
                StreamReader sr = new StreamReader(strFilePath);
                string[] headers = sr.ReadLine().Split(',');
                
                int count = 1;
                foreach (string header in headers)
                {
                    // dt.Columns.Add(header);

                    if (!dt.Columns.Contains(header))
                    {
                        dt.Columns.Add(header);
                    }
                    else
                    {
                        dt.Columns.Add(header + "_" + count);
                        count++;
                    }
                }
                while (!sr.EndOfStream)
                {
                    string[] rows = Regex.Split(sr.ReadLine(), ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < headers.Length; i++)
                    {
                        dr[i] = rows[i];
                    }
                    dt.Rows.Add(dr);
                }
            }
            catch (Exception ex)
            {
                objCommon.WriteErrorLog(ex, "ConvertCSVtoDataSet");
            }
            dsRet.Tables.Add(dt);
            return dsRet;
        }
        public static void ExportDataToXLSX(System.Data.DataTable dt, string strInputFilePath, string fileName)
        {
            clsCommon objCommon = new clsCommon();
            try
            {
                string strFilePath;

                if (!System.IO.Directory.Exists(strInputFilePath + @"\"))
                    System.IO.Directory.CreateDirectory(strInputFilePath + @"\");

                int fileExtPos = fileName.LastIndexOf(".");
                if (fileExtPos >= 0)
                    fileName = fileName.Substring(0, fileExtPos);


                strFilePath = strInputFilePath + @"\" + fileName + ".xlsx"; // ".csv";

                Application oXL;
                Workbook oWB;
                Worksheet oSheet;
                Range oRange;

                try
                {
                    // Start Excel and get Application object. 
                    oXL = new Microsoft.Office.Interop.Excel.Application();

                    // Set some properties 
                    oXL.Visible = false;
                    oXL.DisplayAlerts = false;

                    // Get a new workbook. 
                    oWB = oXL.Workbooks.Add(Type.Missing);

                    // Get the Active sheet 
                    oSheet = (Microsoft.Office.Interop.Excel.Worksheet)oWB.ActiveSheet;
                    oSheet.Name = dt.TableName;

                    //  sda.Fill(dt);
                    //    System.Data.DataTable dt = ds.Tables[0];
                    int rowCount = 1;
                    foreach (DataRow dr in dt.Rows)
                    {
                        rowCount += 1;
                        for (int i = 1; i < dt.Columns.Count + 1; i++)
                        {
                            // Add the header the first time through 
                            if (rowCount == 2)
                            {
                                oSheet.Cells[1, i] = dt.Columns[i - 1].ColumnName;
                            }
                            oSheet.Cells[rowCount, i] = dr[i - 1].ToString();
                        }
                    }

                    // Resize the columns 
                    // Range c1 = oSheet.Cells[1, 1];
                    // Range c2 = oSheet.Cells[rowCount, dt.Columns.Count];
                    //  oRange = oSheet.get_Range(c1, c2);

                    oRange = oSheet.get_Range(oSheet.Cells[1, 1],
                             oSheet.Cells[rowCount, dt.Columns.Count]);

                    oRange.EntireColumn.AutoFit();

                    // Save the sheet and close 
                    oSheet = null;
                    oRange = null;

                    oWB.SaveAs(strFilePath, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
    false, false, XlSaveAsAccessMode.xlNoChange,
    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    oWB.Close(Type.Missing, Type.Missing, Type.Missing);
                    oWB = null;
                    oXL.Quit();
                }
                catch (Exception ex)
                {
                    objCommon.WriteErrorLog(ex, "ExportDataToXLSX");
                    throw;
                }
                finally
                {
                    // Clean up 
                    // NOTE: When in release mode, this does the trick 
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }

            }
            catch (Exception ex)
            {
                objCommon.WriteErrorLog(ex, "ExportDataToXLSX");
            }
        }



    }
}
