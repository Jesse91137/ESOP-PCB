using System;
using System.Data;
using System.Data.SqlClient;

namespace PCB_TO_SOP_DB
{
    /// <summary>
    /// 提供 E_SOP_PCB_Table 資料表的查詢與操作服務。
    /// </summary>
    /// <remarks>
    /// 主要用於判斷資料是否存在、插入新資料、更新現有資料。
    /// </remarks>
    public class PcbRepository
    {
        private readonly string _connStr;

        /// <summary>
        /// 建立 <see cref="PcbRepository"/> 物件，並指定資料庫連線字串。
        /// </summary>
        /// <param name="connStr">SQL Server 連線字串。</param>
        public PcbRepository(string connStr)
        {
            _connStr = connStr;
        }

        /// <summary>
        /// 判斷指定的 Eng_SR 與 PCB_item 是否存在於 E_SOP_PCB_Table。
        /// </summary>
        /// <param name="engSr">工程序號。</param>
        /// <param name="pcbItem">PCB 項目。</param>
        /// <returns>若存在則回傳 <see langword="true"/>，否則回傳 <see langword="false"/>。</returns>
        /// <remarks>
        /// 只檢查是否有符合條件的資料列，不回傳詳細內容。
        /// </remarks>
        /// <example>
        /// <code>
        /// var repo = new PcbRepository(connStr);
        /// bool exists = repo.Exists("SR001", "PCB-A");
        /// </code>
        /// </example>
        public bool Exists(string engSr, string pcbItem)
        {
            string sql = "select 1 from E_SOP_PCB_Table where Eng_SR=@engSr and PCB_item=@pcbItem";
            using (var con = new SqlConnection(_connStr))
            using (var cmd = new SqlCommand(sql, con))
            {
                cmd.Parameters.AddWithValue("@engSr", engSr);
                cmd.Parameters.AddWithValue("@pcbItem", pcbItem);
                con.Open();
                using (var dr = cmd.ExecuteReader())
                    return dr.Read();
            }
        }

        /// <summary>
        /// 新增一筆 Eng_SR 與 PCB_item 至 E_SOP_PCB_Table。
        /// </summary>
        /// <param name="engSr">機種/工程編號。</param>
        /// <param name="pcbItem">PCB 料號。</param>
        /// <remarks>
        /// 若資料已存在，將造成主鍵衝突例外。
        /// </remarks>
        /// <exception cref="SqlException">資料庫寫入失敗時拋出。</exception>
        /// <example>
        /// <code>
        /// var repo = new PcbRepository(connStr);
        /// repo.Insert("SR001", "PCB-A");
        /// </code>
        /// </example>
        public void Insert(string engSr, string pcbItem)
        {
            string sql = "insert into E_SOP_PCB_Table values(@engSr, @pcbItem)";
            ExecuteNonQuery(sql, engSr, pcbItem);
        }

        /// <summary>
        /// 更新指定 Eng_SR 的 PCB_item 欄位。
        /// </summary>
        /// <param name="engSr">機種/工程編號。</param>
        /// <param name="pcbItem">PCB 料號。</param>
        /// <remarks>
        /// 若指定 Eng_SR 不存在，則不會有任何資料被更新。
        /// </remarks>
        /// <exception cref="SqlException">資料庫更新失敗時拋出。</exception>
        /// <example>
        /// <code>
        /// var repo = new PcbRepository(connStr);
        /// repo.Update("SR001", "PCB-B");
        /// </code>
        /// </example>
        public void Update(string engSr, string pcbItem)
        {
            string sql = "update E_SOP_PCB_Table set PCB_item=@pcbItem where Eng_SR=@engSr";
            ExecuteNonQuery(sql, engSr, pcbItem);
        }

        /// <summary>
        /// 執行 SQL 非查詢命令 (Insert/Update)。
        /// </summary>
        /// <param name="sql">要執行的 SQL 指令。</param>
        /// <param name="engSr">機種/工程編號。</param>
        /// <param name="pcbItem">PCB 料號。</param>
        /// <remarks>
        /// 內部方法，僅供 <see cref="Insert"/> 和 <see cref="Update"/> 呼叫。
        /// </remarks>
        /// <exception cref="SqlException">資料庫操作失敗時拋出。</exception>
        private void ExecuteNonQuery(string sql, string engSr, string pcbItem)
        {
            // 使用交易確保資料一致性，並捕捉例外顯示錯誤
            using (var con = new SqlConnection(_connStr))
            {
                con.Open();
                using (var tran = con.BeginTransaction())
                using (var cmd = new SqlCommand(sql, con, tran))
                {
                    cmd.Parameters.AddWithValue("@engSr", engSr);
                    cmd.Parameters.AddWithValue("@pcbItem", pcbItem);
                    try
                    {
                        // 執行 SQL 指令
                        cmd.ExecuteNonQuery();
                        // 提交交易
                        tran.Commit();
                    }
                    catch (Exception ex)
                    {
                        // 發生錯誤時回滾交易並顯示錯誤訊息
                        tran.Rollback();
                        Console.WriteLine($"[資料庫錯誤] {ex.Message}\n{ex.StackTrace}");
                        // 重新拋出例外給上層處理
                        throw;
                    }
                }
            }
        }
    }
}
