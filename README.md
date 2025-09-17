
# PCB_TO_SOP_DB

## 專案簡介

本專案為 PCB 轉 SOP DB 的工具程式，主要功能為將 PCB 資料從 Excel 或其他支援格式轉換並匯入 SOP 資料庫。程式以 C# (.NET Framework 4.8) 開發，整合多個第三方函式庫以處理 Excel、壓縮、影像與加解密等需求。

本文件側重於「系統操作詳細說明」，讓使用者（含系統管理者與開發者）能快速上手系統的安裝、設定、建置、執行與除錯流程。

---

## 目錄結構（重點檔案）

- `Program.cs`：主程式進入點，負責解析執行參數並啟動工作流程
- `ExcelService.cs`：處理 Excel 檔案讀寫與資料驗證
- `PcbRepository.cs`：與 SOP 資料庫互動的資料存取層
- `App.config`：應用程式設定（資料庫連線字串、日誌等）
- `packages.config`：NuGet 相依套件清單
- `bin/`、`obj/`：編譯輸出與中繼檔
- `Properties/AssemblyInfo.cs`：組件資訊

---

## 前置需求

- 作業系統：Windows 10/11
- .NET Framework：4.8
- 開發工具：Visual Studio 2019 或 Visual Studio 2022（若使用命令列則需安裝 MSBuild 與 NuGet CLI）
- 備註：若要在其他機器上執行，請確認目標機器已安裝相容的 .NET Framework 與必要的相依套件。

---

## 安裝與還原相依套件

1. 開啟 Visual Studio，載入 `PCB_TO_SOP_DB.sln`
2. 使用 NuGet 還原相依套件（Visual Studio 會自動還原）。若要在命令列手動執行：

```powershell
# 在專案根目錄下執行（Windows PowerShell）
nuget restore .\PCB_TO_SOP_DB.sln
msbuild .\PCB_TO_SOP_DB.sln /p:Configuration=Release
```

註：若您使用的是 `dotnet` CLI（較新專案），請改用 `dotnet restore` 與 `dotnet build`；但本專案為 .NET Framework，故建議使用 NuGet + MSBuild。

---

## 建置（Build）

- 在 Visual Studio：選擇 Release 或 Debug，按下「建置解決方案」
- 命令列：使用 MSBuild

```powershell
msbuild .\PCB_TO_SOP_DB.sln /p:Configuration=Release
```

建置成功後，執行檔會在 `bin\Release\`（或 `bin\Debug\`）資料夾內。

---

## 執行方式（Run）

本程式可透過 GUI（若有）或命令列參數執行（依 `Program.cs` 的實作而定）。下面提供兩種常見方式：

1. 以 Visual Studio 執行（開發/偵錯） — 在 Visual Studio 中選取要啟動的專案，按 F5（偵錯）或 Ctrl+F5（不偵錯）

1. 以命令列執行（生產或一次性匯入）

```powershell
# 範例：在 bin\Release 下直接執行
cd .\bin\Release
.\PCB_TO_SOP_DB.exe -i "C:\input\pcb_list.xlsx" -c ".\App.config" -l "info"
```

下面的「執行參數」小節說明常見的參數與行為（若專案未實作，請視為建議實作或請求我代為新增）。

---

## 執行參數（假設與建議）

註：程式參數會依 `Program.cs` 實作而異；此處列出常用的參數設計，若您的專案尚未實作這些參數，我可以協助將其加入。

- `-i, --input`：輸入檔案路徑（Excel）。例如：`-i "C:\input\pcb_list.xlsx"`
- `-m, --mode`：執行模式（例如：`preview`、`import`）。`preview` 僅做驗證與報表；`import` 真正匯入資料庫。
- `-c, --config`：指定設定檔路徑（預設使用同目錄下的 `App.config`）
- `-l, --log`：日誌等級（`debug`、`info`、`warn`、`error`）
- `-o, --output`：輸出報表或錯誤檔案路徑

範例執行：

```powershell
.\PCB_TO_SOP_DB.exe -i "C:\input\pcb_list.xlsx" -m import -l info
```

假設：若您希望我直接為 `Program.cs` 加入上述參數解析（使用 `CommandLineParser` 或手寫 `args` 分析），請回覆同意，我會進入實作模式。

---

## 設定檔（App.config）重點說明

請檢查 `App.config` 中下列常見設定項：

- 資料庫連線字串（ConnectionStrings）——必須填入正確的伺服器、資料庫、授權方式
- 日誌路徑／級別設定（若專案使用 log 框架）
- 匯入選項（例如是否允許覆寫既有料號、批次大小）

範例（說明性，請以專案實際內容為準）：

```xml
<!-- App.config 範例片段 -->
<configuration>
  <connectionStrings>
    <add name="SopDb" connectionString="Server=SERVER;Database=SOP;User Id=sa;Password=secret;" />
  </connectionStrings>
  <!-- 其他設定... -->
</configuration>
```

若您不確定 `App.config` 中每個鍵的用途，我可以協助掃描專案並列出該設定的使用位置與預期格式。

---

## 日誌與錯誤處理

- 程式執行期間會輸出日誌，請確認 `App.config` 或專案中的日誌設定（例如檔案路徑、輪替機制）可寫入
- 若發生錯誤，先查看日誌（Log）與輸出報表（若有）以取得詳細錯誤訊息
- 常見錯誤原因：Excel 欄位缺失、資料型別不符、資料庫連線失敗、權限不足

除錯步驟建議：

1. 以 `preview` 或等價的驗證模式先執行，確認資料格式與映射正確
2. 檢查 `App.config` 中的連線字串、帳號密碼、網路連線
3. 在本地端以 Visual Studio 偵錯，打斷點觀察 `ExcelService` 與 `PcbRepository` 的行為

---

## 常見問題（FAQ）與排除建議

- 問：無法還原 NuGet 套件？
  - 建議：確認網路連線與 NuGet 源設定；可手動於專案資料夾執行 `nuget restore`。
- 問：程式執行時找不到某個 DLL？
  - 建議：確認 `bin` 目錄下的相依套件是否完整，或重新建置解決方案。
- 問：資料庫匯入失敗但沒有詳細錯誤？
  - 建議：提升日誌等級至 `debug` 並重新執行以取得更多錯誤細節；或用 SQL Server Profiler/查詢記錄檢查伺服器端錯誤。

---

## 開發者提示（建議的改善小項目）

- 在 `Program.cs` 加入明確的 CLI 參數解析與 `--help` 說明，提升操作友善性
- 支援匯入前的欄位對照（mapping）預覽報表，讓使用者先確認欄位對應
- 將日誌統一改由成熟的記錄函式庫（如 NLog、Serilog）並支援滾動檔與遠端收集

---

## 嘗試執行（範例流程）

1. 還原套件與建置：

```powershell
nuget restore .\PCB_TO_SOP_DB.sln; msbuild .\PCB_TO_SOP_DB.sln /p:Configuration=Release
```

1. 在 Release 資料夾執行預覽（假設提供 preview 模式）：

```powershell
cd .\bin\Release
.\PCB_TO_SOP_DB.exe -i "C:\input\pcb_list.xlsx" -m preview -o "C:\output\report.csv" -l debug
```

1. 若預覽正確，執行匯入：

```powershell
.\PCB_TO_SOP_DB.exe -i "C:\input\pcb_list.xlsx" -m import -l info
```

---

## 聯絡與回報問題

如有問題或建議，請提交 Issue 或聯絡專案負責人。提交 Issue 時請附上：

- 錯誤日誌（若有）
- 測試用的 Excel 範例（去除敏感資訊）
- `App.config` 中的相關設定（連線字串可用占位符代替）

---

> **免責聲明**：本文件為專案使用及操作說明，內容可依實際程式實作而調整；若 README 中某些執行參數或行為與 `Program.cs` 不符，我可以協助掃描程式並同步更新說明。


