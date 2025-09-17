using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PCB_TO_SOP_DB
{
    /// <summary>
    /// 提供 Excel 檔案與 DataTable 之間轉換的服務。
    /// </summary>
    /// <remarks>
    /// 使用 NPOI 函式庫處理 .xls 與 .xlsx 格式。
    /// </remarks>
    internal class ExcelService
    {
        /// <summary>
        /// 讀取 Excel 檔案並轉換為 <see cref="DataTable"/>。
        /// </summary>
        /// <param name="xlsFilename">Excel 檔案路徑。</param>
        /// <returns>轉換後的 <see cref="DataTable"/>。</returns>
        /// <remarks>
        /// 只讀取第一個工作表，並假設第一列為標題列。
        /// 儲存格型別為數值時存入 double，否則存入 string。
        /// </remarks>
        /// <exception cref="FileNotFoundException">檔案不存在時拋出。</exception>
        /// <example>
        /// <code>
        /// var service = new ExcelService();
        /// DataTable dt = service.LoadExcelAsDataTable("data.xlsx");
        /// </code>
        /// </example>
        public DataTable LoadExcelAsDataTable(string xlsFilename)
        {
            FileInfo fi = new FileInfo(xlsFilename); // 取得檔案資訊
            using (FileStream fstream = new FileStream(fi.FullName, FileMode.Open)) // 開啟檔案串流
            {
                IWorkbook wb; // 宣告 Excel 工作簿物件
                if (fi.Extension == ".xlsx") // 判斷副檔名是否為 .xlsx
                    wb = new XSSFWorkbook(fstream); // 建立 XSSFWorkbook 物件 - Excel 2007
                else
                    wb = new HSSFWorkbook(fstream); // 建立 HSSFWorkbook 物件 - Excel 97-2003

                ISheet sheet = wb.GetSheetAt(0); // 取得第一個工作表(工程部-寶雅(電腦-DSP47)製好後會把PCB資料移到第一張工作表)
                DataTable table = new DataTable(); // 建立 DataTable 物件
                IRow headerRow = sheet.GetRow(0); // 取得標題列
                int cellCount = headerRow.LastCellNum; // 取得欄位數
                for (int i = headerRow.FirstCellNum; i < cellCount; i++) // 逐欄建立 DataTable 欄位
                    table.Columns.Add(new DataColumn(headerRow.GetCell(i).StringCellValue)); // 新增欄位名稱

                for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++) // 逐行讀取資料
                {
                    IRow row = sheet.GetRow(i); // 取得目前資料列
                    if (row == null) continue; // 若資料列為空則跳過
                    DataRow dataRow = table.NewRow(); // 建立新的 DataRow
                    for (int j = row.FirstCellNum; j < cellCount; j++) // 逐欄讀取儲存格
                    {
                        ICell cell = row.GetCell(j); // 取得儲存格物件
                        if (cell != null) // 若儲存格不為空
                        {
                            // 若儲存格型別為數值，則存入 double，否則存入 string
                            if (cell.CellType == CellType.Numeric) // 判斷是否為數值型別
                                dataRow[j] = cell.NumericCellValue; // 存入 double
                            else
                                dataRow[j] = cell.ToString(); // 存入字串
                        }
                    }
                    table.Rows.Add(dataRow); // 將資料列加入 DataTable
                }
                return table; // 回傳 DataTable
            }
        }

        /// <summary>
        /// 將 <see cref="DataTable"/> 儲存為 Excel 檔案。
        /// </summary>
        /// <param name="table">要儲存的 <see cref="DataTable"/>。</param>
        /// <param name="xlsFilename">儲存的 Excel 檔案路徑。</param>
        /// <param name="isXlsx">是否儲存為 .xlsx 格式，否則為 .xls。</param>
        /// <remarks>
        /// 只會產生一個名為 "Sheet1" 的工作表，且所有資料皆以字串型別儲存。
        /// </remarks>
        /// <exception cref="IOException">檔案寫入失敗時拋出。</exception>
        /// <example>
        /// <code>
        /// var service = new ExcelService();
        /// service.SaveDataTableToExcel(dt, "output.xlsx", true);
        /// </code>
        /// </example>
        public void SaveDataTableToExcel(DataTable table, string xlsFilename, bool isXlsx)
        {
            object workbook; // 宣告工作簿物件
            if (isXlsx) // 判斷是否為 .xlsx 格式
                workbook = new XSSFWorkbook(); // 建立 XSSFWorkbook 物件
            else
                workbook = new HSSFWorkbook(); // 建立 HSSFWorkbook 物件

            ISheet sheet = ((IWorkbook)workbook).CreateSheet("Sheet1"); // 建立名為 Sheet1 的工作表
            IRow headerRow = sheet.CreateRow(0); // 建立標題列
            for (int i = 0; i < table.Columns.Count; i++) // 逐欄寫入欄位名稱
                headerRow.CreateCell(i).SetCellValue(table.Columns[i].ColumnName); // 設定欄位名稱

            for (int i = 0; i < table.Rows.Count; i++) // 逐行寫入資料
            {
                IRow row = sheet.CreateRow(i + 1); // 建立資料列
                for (int j = 0; j < table.Columns.Count; j++) // 逐欄寫入儲存格
                {
                    row.CreateCell(j).SetCellValue(table.Rows[i][j].ToString()); // 設定儲存格內容
                }
            }

            using (FileStream fs = new FileStream(xlsFilename, FileMode.Create)) // 建立檔案串流
            {
                ((IWorkbook)workbook).Write(fs); // 將工作簿寫入檔案
            }
        }
    }
}
