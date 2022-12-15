using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
using Excel = Microsoft.Office.Interop.Excel;
using ADODB;
namespace JWLRetriveEmail
{
    class clsExcelHelperNew : clsCommon
    {

        public static ReturnResponse ExportDatasetToXLSX(DataSet ds, string strOutputFilePath, string fileName)
        {
            ReturnResponse objReturnResponse = new ReturnResponse();
            clsCommon objCommon = new clsCommon();
            try
            {
                string outputFileLocation;
                string outputFile;

                outputFileLocation = strOutputFilePath;// + @"\Outputs";

                if (!System.IO.Directory.Exists(outputFileLocation + @"\"))
                    System.IO.Directory.CreateDirectory(outputFileLocation + @"\");


                int fileExtPos = fileName.LastIndexOf(".");
                if (fileExtPos >= 0)
                    fileName = fileName.Substring(0, fileExtPos);

                outputFile = fileName;
                outputFile = outputFileLocation + @"\" + outputFile + ".xlsx"; // ".csv";

                if (File.Exists(outputFile))
                {
                    File.Delete(outputFile);
                }

                Excel.Application m_objExcel = null;
                Excel.Workbooks m_objBooks = null;
                Excel._Workbook m_objBook = null;
                Excel.Sheets m_objSheets = null;
                Excel._Worksheet m_objSheet = null;
                Excel.Range m_objRange = null;
                Excel.Font m_objFont = null;
               
                object m_objOpt = System.Reflection.Missing.Value;

                m_objExcel = new Excel.Application();
                m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
                m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));
                m_objSheets = (Excel.Sheets)m_objBook.Worksheets;

                DataTableCollection collection = ds.Tables;
                for (int i = collection.Count; i > 0; i--)
                {
                    ADODB._Recordset objRS = null;
                  
                    m_objSheet = null;

                    // m_objSheet = (Excel._Worksheet)(m_objSheets.get_Item(sheetnumber+1));
                    m_objSheet = (Excel._Worksheet)m_objSheets.Add(m_objSheets[1],
                                  Type.Missing, Type.Missing, Type.Missing);

                    System.Data.DataTable table = collection[i - 1];
                    m_objSheet.Name = table.TableName;

                    objRS = ConvertToRecordset(table);
                    int nFields = table.Columns.Count;

                    // Create an array for the headers and add it to the
                    // worksheet starting at cell A1.
                    object[] objHeaders = new object[nFields];
                    for (int j = 1; j < table.Columns.Count + 1; j++)
                    {
                        objHeaders[j - 1] = table.Columns[j - 1].ColumnName;
                    }

                    m_objRange = m_objSheet.get_Range("A1", m_objOpt);
                    m_objRange = m_objRange.get_Resize(1, nFields);
                    m_objRange.set_Value(m_objOpt, objHeaders);
                    m_objFont = m_objRange.Font;
                    m_objFont.Bold = true;

                    // Transfer the recordset to the worksheet starting at cell A2.
                    m_objRange = m_objSheet.get_Range("A2", m_objOpt);
                    m_objRange.CopyFromRecordset(objRS, m_objOpt, m_objOpt);
                }



                // Save the workbook and quit Excel.
                m_objBook.SaveAs(outputFile, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                   m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                m_objBook.Close(false, m_objOpt, m_objOpt);
                m_objExcel.Quit();

                objReturnResponse.ResponseVal = true;
            }
            catch (Exception ex)
            {
                string strExecutionLogMessage = "Exception in ExportDatasetToXLSX" + System.Environment.NewLine;
                strExecutionLogMessage += "fileName - " + fileName + System.Environment.NewLine;
                objCommon.WriteErrorLog(ex, strExecutionLogMessage);
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
            return objReturnResponse;
        }

        public static ReturnResponse ExportDataTableToXLSX(System.Data.DataTable dt, string strOutputFilePath, string fileName)
        {
            ReturnResponse objReturnResponse = new ReturnResponse();
            clsCommon objCommon = new clsCommon();
            try
            {
                string outputFileLocation;
                string outputFile;

                outputFileLocation = strOutputFilePath;// + @"\Outputs";

                if (!System.IO.Directory.Exists(outputFileLocation + @"\"))
                    System.IO.Directory.CreateDirectory(outputFileLocation + @"\");


                int fileExtPos = fileName.LastIndexOf(".");
                if (fileExtPos >= 0)
                    fileName = fileName.Substring(0, fileExtPos);

                outputFile = fileName;
                outputFile = outputFileLocation + @"\" + outputFile + ".xlsx"; // ".csv";

                if (File.Exists(outputFile))
                {
                    File.Delete(outputFile);
                }

                Excel.Application m_objExcel = null;
                Excel.Workbooks m_objBooks = null;
                Excel._Workbook m_objBook = null;
                Excel.Sheets m_objSheets = null;
                Excel._Worksheet m_objSheet = null;
                Excel.Range m_objRange = null;
                Excel.Font m_objFont = null;
                object m_objOpt = System.Reflection.Missing.Value;

                m_objExcel = new Excel.Application();
                m_objBooks = (Excel.Workbooks)m_objExcel.Workbooks;
                m_objBook = (Excel._Workbook)(m_objBooks.Add(m_objOpt));
                m_objSheets = (Excel.Sheets)m_objBook.Worksheets;

                ADODB._Recordset objRS = null;


                m_objSheet = (Excel._Worksheet)m_objSheets.Add(m_objSheets[1],
                              Type.Missing, Type.Missing, Type.Missing);

                System.Data.DataTable table = dt;
                m_objSheet.Name = table.TableName;


                objRS = ConvertToRecordset(table);
                int nFields = table.Columns.Count;

                // Create an array for the headers and add it to the
                // worksheet starting at cell A1.
                object[] objHeaders = new object[nFields];
                for (int j = 1; j < table.Columns.Count + 1; j++)
                {
                    objHeaders[j - 1] = table.Columns[j - 1].ColumnName;
                }

                m_objRange = m_objSheet.get_Range("A1", m_objOpt);
                m_objRange = m_objRange.get_Resize(1, nFields);
                m_objRange.set_Value(m_objOpt, objHeaders);
                m_objFont = m_objRange.Font;
                m_objFont.Bold = true;

                // Transfer the recordset to the worksheet starting at cell A2.
                m_objRange = m_objSheet.get_Range("A2", m_objOpt);
                m_objRange.CopyFromRecordset(objRS, m_objOpt, m_objOpt);


                // Save the workbook and quit Excel.
                m_objBook.SaveAs(outputFile, m_objOpt, m_objOpt,
                    m_objOpt, m_objOpt, m_objOpt, Excel.XlSaveAsAccessMode.xlNoChange,
                   m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                m_objBook.Close(false, m_objOpt, m_objOpt);
                m_objExcel.Quit();

                objReturnResponse.ResponseVal = true;
            }
            catch (Exception ex)
            {
                string strExecutionLogMessage = "Exception in ExportDataTableToXLSX" + System.Environment.NewLine;
                strExecutionLogMessage += "fileName - " + fileName + System.Environment.NewLine;
                objCommon.WriteErrorLog(ex, strExecutionLogMessage);
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
            return objReturnResponse;
        }
        static public ADODB.Recordset ConvertToRecordset(System.Data.DataTable inTable)
        {
            ADODB.Recordset result = new ADODB.Recordset();
            result.CursorLocation = ADODB.CursorLocationEnum.adUseClient;

            ADODB.Fields resultFields = result.Fields;
            System.Data.DataColumnCollection inColumns = inTable.Columns;

            foreach (DataColumn inColumn in inColumns)
            {
                resultFields.Append(inColumn.ColumnName
                    , TranslateType(inColumn.DataType)
                    , inColumn.MaxLength
                    , inColumn.AllowDBNull ? ADODB.FieldAttributeEnum.adFldIsNullable :
                                             ADODB.FieldAttributeEnum.adFldUnspecified
                    , null);
            }

            result.Open(System.Reflection.Missing.Value
                    , System.Reflection.Missing.Value
                    , ADODB.CursorTypeEnum.adOpenStatic
                    , ADODB.LockTypeEnum.adLockOptimistic, 0);

            foreach (DataRow dr in inTable.Rows)
            {
                result.AddNew(System.Reflection.Missing.Value,
                              System.Reflection.Missing.Value);

                for (int columnIndex = 0; columnIndex < inColumns.Count; columnIndex++)
                {
                    resultFields[columnIndex].Value = dr[columnIndex];
                }
            }

            return result;
        }

        static ADODB.DataTypeEnum TranslateType(Type columnType)
        {
            switch (columnType.UnderlyingSystemType.ToString())
            {
                case "System.Boolean":
                    return ADODB.DataTypeEnum.adBoolean;

                case "System.Byte":
                    return ADODB.DataTypeEnum.adUnsignedTinyInt;

                case "System.Char":
                    return ADODB.DataTypeEnum.adChar;

                case "System.DateTime":
                    return ADODB.DataTypeEnum.adDate;

                case "System.Decimal":
                    return ADODB.DataTypeEnum.adCurrency;

                case "System.Double":
                    return ADODB.DataTypeEnum.adDouble;

                case "System.Int16":
                    return ADODB.DataTypeEnum.adSmallInt;

                case "System.Int32":
                    return ADODB.DataTypeEnum.adInteger;

                case "System.Int64":
                    return ADODB.DataTypeEnum.adBigInt;

                case "System.SByte":
                    return ADODB.DataTypeEnum.adTinyInt;

                case "System.Single":
                    return ADODB.DataTypeEnum.adSingle;

                case "System.UInt16":
                    return ADODB.DataTypeEnum.adUnsignedSmallInt;

                case "System.UInt32":
                    return ADODB.DataTypeEnum.adUnsignedInt;

                case "System.UInt64":
                    return ADODB.DataTypeEnum.adUnsignedBigInt;

                case "System.String":
                default:
                    return ADODB.DataTypeEnum.adVarChar;
            }
        }

    }
}
