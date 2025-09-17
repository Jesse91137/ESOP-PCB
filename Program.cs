using System;
using System.Configuration;
using System.Data;
using System.IO;

namespace PCB_TO_SOP_DB
{
    /// <summary>
    /// E-SOP
    /// PCB料號、機種 Excel檔案寫入DB
    /// </summary>
    /// <remarks>
    /// 讀取 Excel 資料，將資料寫入 E_SOP_PCB_Table 資料表。
    /// </remarks>
    internal class Program
    {
        /// <summary>
        /// 主程式進入點。
        /// </summary>
        /// <param name="args">命令列參數。</param>
        /// <example>
        /// <code>
        /// // 執行方式
        /// Program.exe
        /// </code>
        /// </example>
        static void Main(string[] args)
        {
            try
            {
                // 在控制台輸出 "資料寫入中....."
                Console.WriteLine("資料寫入中.....");
                // 取得 Excel 檔案資料夾路徑
                string xlsPath = ConfigurationManager.AppSettings["XLS_PATH"];
                // 取得資料庫連線字串
                string connStr = ConfigurationManager.AppSettings["CNN_TEXT"];
                // 建立 Excel 服務物件
                var excelService = new ExcelService();
                // 建立資料庫操作物件
                var repo = new PcbRepository(connStr);

                // 讀取指定路徑的資料夾
                DirectoryInfo xlsDir = new DirectoryInfo(xlsPath);
                // 若資料夾不存在則結束程式
                if (!xlsDir.Exists) return;

                // 逐一處理取得目錄下所有 .xlsx 檔案
                foreach (var xlsFile in xlsDir.GetFiles("*.xlsx"))
                {
                    // 讀取 Excel 檔案為 DataTable
                    DataTable dt = excelService.LoadExcelAsDataTable(xlsFile.FullName);
                    // 逐行處理資料
                    foreach (DataRow row in dt.Rows)
                    {
                        // 取得機種/工程編號並轉大寫
                        string engSr = row[0].ToString().ToUpper();
                        // 取得 PCB 料號
                        string pcbItem = row[2].ToString();
                        // 判斷資料是否存在
                        if (!repo.Exists(engSr, pcbItem))
                        {
                            // 若不存在則新增資料
                            repo.Insert(engSr, pcbItem);
                            // 顯示新增訊息
                            Console.WriteLine($"\n 新增 {engSr} , {pcbItem}");
                        }
                        else
                        {
                            // 若已存在則更新資料
                            repo.Update(engSr, pcbItem);
                            // 顯示更新訊息
                            Console.WriteLine($"\n 更新 {engSr} , {pcbItem}");
                        }
                    }
                }
                // 顯示完成訊息
                Console.WriteLine("\n\n\n\n寫入完畢,按任意建關閉!!");
            }
            catch (Exception ex)
            {
                // 顯示錯誤訊息與堆疊資訊
                Console.WriteLine($"\n[錯誤] {ex.Message}\n{ex.StackTrace}");
            }
            finally
            {
                // 等待使用者按鍵後結束
                Console.WriteLine("\n按任意鍵關閉程式...");
                Console.ReadKey();
            }
        }
    }
}
