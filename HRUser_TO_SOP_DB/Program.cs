using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PCB_TO_SOP_DB
{
    class Program
    {
        private static string xlsPath = "";
        private static string cnnText = "";
        private static readonly String connStr = ConfigurationManager.AppSettings["CNN_TEXT"].ToString();
        static void Main(string[] args)
        {
            Console.WriteLine("資料寫入中.....");
            try
            {
                xlsPath = ConfigurationManager.AppSettings["XLS_PATH"].ToString();
            }
            catch { }
            try
            {
                cnnText = ConfigurationManager.AppSettings["CNN_TEXT"].ToString();
            }
            catch { }
            DirectoryInfo xlsDir = new DirectoryInfo(xlsPath);
            if (xlsDir.Exists)
            {
                FileInfo[] xlsFiles = xlsDir.GetFiles("*.xlsx");
                DataTable dt = new DataTable();
                try
                {
                    for (int i = 0; i < xlsFiles.Length; i++)
                    {
                        string xlsFileName = xlsFiles[i].FullName;
                        dt = LoadExcelAsDataTable(xlsFileName);
                        #region MS Excel Method                        
                        //string xlsFileName = xlsFiles[i].FullName;
                        //string excelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + xlsFileName + ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
                        //OleDbConnection cnn = new OleDbConnection(excelString);
                        //OleDbCommand cmd = new OleDbCommand();
                        //cmd.Connection = cnn;
                        //OleDbDataAdapter adapter = new OleDbDataAdapter();
                        //cmd.CommandText = "SELECT * FROM [在職人員一覽表$]";
                        //adapter.SelectCommand = cmd;
                        //DataSet dsData = new DataSet();
                        //adapter.Fill(dsData);
                        //cnn = null;
                        //cmd = null;
                        //adapter = null;
                        //string errorDesc = "";
                        #endregion
                        //数据从第2行开始
                        foreach (DataRow row in dt.Rows)
                        {
                            string engSr = row[0].ToString().ToUpper();
                            string pcbItem = row[2].ToString();                            

                            string sql = "select * from E_SOP_PCB_Table where Eng_SR=@engSr and PCB_item=@pcbItem ";
                            SqlParameter[] pms = new SqlParameter[]
                            {
                                new SqlParameter("@engSr",engSr),
                                new SqlParameter("@pcbItem",pcbItem)                             
                            };
                            SqlDataReader dr = ExecuteReader(sql, CommandType.Text, pms);
                            if (!dr.Read())
                            {
                                string insSql = "insert into E_SOP_PCB_Table values( @engSr  ,@pcbItem) ";
                                SqlParameter[] parm2 = new SqlParameter[]
                                {
                                    new SqlParameter("@engSr",engSr),
                                    new SqlParameter("@pcbItem",pcbItem)
                                };
                                ExecueNonQuery(insSql, CommandType.Text, parm2);
                                Console.WriteLine("\n 新增" + engSr + "  , " + pcbItem);
                            }
                            else
                            {
                                string upSql = "update E_SOP_PCB_Table set PCB_item=@pcbItem where Eng_SR=@engSr ";
                                SqlParameter[] parm2 = new SqlParameter[]
                                {
                                    new SqlParameter("@pcbItem",pcbItem),
                                    new SqlParameter("@engSr",engSr),
                                };
                                ExecueNonQuery(upSql, CommandType.Text, parm2);
                                Console.WriteLine("\n 更新" + engSr + "  , " + pcbItem);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Console.ReadKey();
                }
                #region Delete
                //DirectoryInfo xlsfi = new DirectoryInfo(xlsPath);
                //if (xlsfi.Exists)
                //{
                //    foreach (var fi in xlsfi.GetFiles())
                //    {
                //        File.Delete(fi.FullName);
                //    }
                //}
                #endregion

                Console.WriteLine("\n\n\n\n" + "寫入完畢,按任意建關閉!!");
                Console.ReadKey();


            }
        }
        
        public static int ExecueNonQuery(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            using (SqlConnection con = new SqlConnection(connStr))
            {
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    //設置目前執行的是「存儲過程? 還是帶參數的sql 語句?」
                    cmd.CommandType = cmdType;
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }

                    con.Open();
                    return cmd.ExecuteNonQuery();
                }
            }
        }
        public static SqlDataReader ExecuteReader(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            SqlConnection con = new SqlConnection(connStr);
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.CommandType = cmdType;
                if (pms != null)
                {
                    cmd.Parameters.AddRange(pms);
                }
                try
                {
                    con.Open();
                    return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
                catch
                {
                    con.Close();
                    con.Dispose();
                    throw;
                }
            }
        }
        public static DataTable LoadExcelAsDataTable(String xlsFilename)
        {
            FileInfo fi = new FileInfo(xlsFilename);
            using (FileStream fstream = new FileStream(fi.FullName, FileMode.Open))
            {
                IWorkbook wb;
                if (fi.Extension == ".xlsx")
                    wb = new XSSFWorkbook(fstream); // excel2007
                else
                    wb = new HSSFWorkbook(fstream); // excel97

                // 只取第一個sheet。
                ISheet sheet = wb.GetSheetAt(0);

                // target
                DataTable table = new DataTable();

                // 由第一列取標題做為欄位名稱
                IRow headerRow = sheet.GetRow(0);
                int cellCount = headerRow.LastCellNum; // 取欄位數
                for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                {
                    //table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue, typeof(double)));
                    table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue));
                }

                // 略過第零列(標題列)，一直處理至最後一列
                for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)                
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    DataRow dataRow = table.NewRow();

                    //依先前取得的欄位數逐一設定欄位內容
                    for (int j = row.FirstCellNum; j < cellCount; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            //如要針對不同型別做個別處理，可善用.CellType判斷型別
                            //再用.StringCellValue, .DateCellValue, .NumericCellValue...取值

                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    dataRow[j] = cell.NumericCellValue;
                                    break;
                                default: // String
                                         //此處只簡單轉成字串
                                    dataRow[j] = cell.StringCellValue;
                                    break;
                            }
                        }
                    }

                    table.Rows.Add(dataRow);
                }

                // success
                return table;
            }
        }
    }
}
