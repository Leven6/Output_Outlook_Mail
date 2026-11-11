# Outlook 郵件匯出工具

將指定寄件者的 Outlook 郵件匯出成 `.msg` 檔，保留郵件內容與附件。工具只讀取及匯出郵件，不會移動或刪除 Outlook 裡的原始郵件。

## 使用前準備

- 使用 Windows，並已安裝支援 COM 自動化的傳統版 Outlook 桌面程式。新版 Outlook、網頁版 Outlook 不適用。
- 已在傳統版 Outlook 設定帳號，且能正常開啟收件匣。建議先開啟 Outlook，完成登入及郵件同步。
- 電腦可執行 Windows PowerShell；BAT 使用系統的 `powershell.exe`。
- 將 [Run-OutlookExport.bat](Run-OutlookExport.bat) 與 [Export-OutlookSender.ps1](Export-OutlookSender.ps1) 放在同一個資料夾，並確認輸出位置有寫入權限與足夠空間。

## 方法一：雙擊執行

1. 雙擊 [Run-OutlookExport.bat](Run-OutlookExport.bat)。
2. 看到 `Sender email address:` 時，輸入寄件者的完整電子郵件地址，例如 `sender@example.com`，再按 Enter。不要加引號。
3. 看到 `Start date` 時，輸入開始日期，例如 `2026-10-01`，再按 Enter。
4. 看到 `End date` 時，輸入結束日期，例如 `2026-10-27`，再按 Enter。格式必須是 `YYYY-MM-DD`，不必加引號。
5. 等待匯出完成。這個方式會掃描預設收件匣及其所有子資料夾。
6. 看到 `Done.` 後，查看 BAT 同層的 `outlook-export` 資料夾。視窗會暫停，按任意鍵關閉。

例如依序輸入：

```text
Sender email address: sender@example.com
Start date (blank = no lower limit): 2026-10-01
End date (blank = now; includes the whole day): 2026-10-27
```

這會匯出寄件時間從 **2026-10-01 00:00:00 至 2026-10-27 當天結束**的郵件，包含起訖兩天。開始與結束日期可以填同一天。

- 只填開始日期：從該日零點匯出到本次啟動時間。
- 只填結束日期：從最早的郵件匯出到該日結束。
- 兩個日期都直接按 Enter 留白：使用原本的增量匯出方式，從最新已匯出時間接續匯出。

日期無效或開始日期晚於結束時間時，會顯示錯誤並停止，不會連線 Outlook 或開始匯出。若只填開始日期，該日期也不能晚於目前時間。

輸出資料夾範例：

```text
Mail/
  Export-OutlookSender.ps1
  Run-OutlookExport.bat
  README.md
  outlook-export/
    sender@example.com/
      20261027-093000 - Meeting notes - A1B2C3D4E5.msg
```

不同寄件者會存入各自的資料夾；同一寄件者在不同 Outlook 子資料夾中的郵件，會集中輸出，不會重建原本的資料夾階層。

## 方法二：用 PowerShell 執行

在檔案總管開啟 `Mail` 資料夾，於網址列輸入 `powershell` 並按 Enter，即可在此資料夾開啟 PowerShell。以下命令都假設目前位置是 `Mail`。

**指定日期區間**，包含收件匣的子資料夾：

```powershell
.\Export-OutlookSender.ps1 -Sender 'sender@example.com' -StartDate '2026-10-01' -EndDate '2026-10-27' -Recurse
```

只指定開始日期，匯出到本次啟動時間：

```powershell
.\Export-OutlookSender.ps1 -Sender 'sender@example.com' -StartDate '2026-10-01' -Recurse
```

只指定結束日期，匯出截至該日的所有符合郵件：

```powershell
.\Export-OutlookSender.ps1 -Sender 'sender@example.com' -EndDate '2026-10-27' -Recurse
```

不指定日期，使用增量匯出，包含收件匣的子資料夾：

```powershell
.\Export-OutlookSender.ps1 -Sender 'sender@example.com' -Recurse
```

指定輸出位置：

```powershell
.\Export-OutlookSender.ps1 -Sender 'sender@example.com' -OutputPath 'D:\MailBackup' -Recurse
```

只掃描收件匣本身，不包含子資料夾：

```powershell
.\Export-OutlookSender.ps1 -Sender 'sender@example.com'
```

預覽哪些郵件會被匯出，並顯示詳細掃描訊息：

```powershell
.\Export-OutlookSender.ps1 -Sender 'sender@example.com' -StartDate '2026-10-01' -EndDate '2026-10-27' -Recurse -WhatIf -Verbose
```

`-WhatIf` 不會儲存郵件，但仍會建立輸出資料夾，並連線 Outlook 掃描郵件。預覽會使用指定的日期區間；未指定日期時，則套用下方的增量匯出規則。

### 參數

| 參數 | 用途 | 預設值 |
| --- | --- | --- |
| `-Sender` | 必填，寄件者的完整電子郵件地址；不分大小寫精確比對，不支援顯示名稱、部分地址或萬用字元。 | 無 |
| `-OutputPath` | 輸出根目錄，工具會在其下建立寄件者資料夾。含空白的路徑請加引號。 | PowerShell 目前工作目錄下的 `outlook-export` |
| `-Recurse` | 同時掃描收件匣的所有子資料夾。 | 不包含子資料夾；BAT 會自動加上此參數 |
| `-StartDate` | 開始日期，格式 `yyyy-MM-dd`，包含當日零點起的郵件。 | 單獨省略時無下限；起訖皆省略時使用增量模式 |
| `-EndDate` | 結束日期，格式 `yyyy-MM-dd`，包含當天整日。 | 本次啟動時間 |
| `-WhatIf` | 預覽匯出動作，不儲存 `.msg` 檔。 | 關閉 |
| `-Verbose` | 顯示掃描資料夾及匯出檔案的詳細訊息。 | 關閉 |

BAT 固定輸出到 BAT 所在資料夾下的 `outlook-export`；直接執行 PowerShell 腳本時，預設位置則取決於目前工作目錄。

## 匯出範圍與增量規則

- **資料夾範圍：**只掃描 Outlook 目前設定檔的預設收件匣；加上 `-Recurse` 才包含其子資料夾。不會自動掃描其他帳號、共用信箱、封存資料夾、寄件備份或整個信箱。
- **時間依據：**使用郵件的寄件時間 `SentOn`，不是收件時間；日期依電腦本地時間解讀，不接受時分秒或時區格式。未指定結束日期時，上限是本次啟動匯出時的電腦時間。
- **手動日期區間：**只要填寫任一日期，就忽略現有匯出檔的最新時間，完整檢查指定區間。已存在相同目標檔名仍跳過，不覆寫。
- **增量模式第一次執行：**起訖日期皆留白，且寄件者輸出資料夾內沒有可辨識的 `.msg` 檔名時，匯出掃描範圍內、截至啟動時間的所有符合郵件。
- **增量模式後續執行：**起訖日期皆留白時，讀取寄件者資料夾內檔名中的最新時間，只處理該時間起至本次啟動時間的郵件。時間相同的郵件仍會檢查，已存在相同目標檔名則跳過，不覆寫。
- **檔名格式：**`yyyyMMdd-HHmmss - 郵件主旨 - 識別碼.msg`。識別碼由 Outlook 的 `EntryID` 計算；主旨中的非法檔名字元會替換成 `_`，主旨最多保留 100 個字元，空白主旨使用 `no-subject`。
- **附件處理：**附件保留在 `.msg` 郵件內，不會另外拆成獨立檔案。可使用傳統版 Outlook 開啟 `.msg` 檢視郵件與附件。

**注意：增量匯出不是完整的缺檔補齊。** 若舊郵件晚同步進來、先前匯出失敗、刪除了較舊的匯出檔，或後來才開啟子資料夾掃描，只要郵件寄件時間早於現有最新匯出時間，就不會被補匯。

需要補匯舊郵件時，可直接指定包含那些郵件的日期區間；既有檔案不會覆寫。需要所有郵件重新輸出成新檔時，請用 `-OutputPath` 指向新的空白輸出根目錄，並省略日期或指定涵蓋全部郵件的區間。不要混用不同信箱的匯出結果，也不要手動改動用來判斷增量時間的檔名前綴。

## 完成訊息

```text
Done. Exported: 12; already present: 3; failed: 0.
Output: ...\outlook-export\sender@example.com
```

- `Exported`：本次成功匯出的郵件數。
- `already present`：本次處理時間範圍內，因相同目標檔名已存在而跳過的數量；不是所有歷次匯出的總數。
- `failed`：本次處理單封郵件時發生錯誤的數量，原因會顯示為警告。這個數字不一定只包含目標寄件者的郵件。
- `Output`：實際的寄件者輸出資料夾。

單封郵件失敗時，腳本會顯示警告並繼續；因此即使出現 `Done.`，也應確認 `failed` 是否為 `0`。Outlook 連線或其他未處理的錯誤則可能中斷整次執行。

## 常見問題

### 沒有匯出任何郵件

確認寄件者地址正確、郵件位於預設收件匣或已納入掃描的子資料夾，並等待 Outlook 同步完成。確認所填區間包含郵件的寄件日期；如果未填日期且曾經匯出過，請同時檢查增量時間範圍。可加上 `-Verbose` 或用 `-WhatIf -Verbose` 查看掃描情況。

### 日期格式錯誤或起訖順序錯誤

請使用 `2026-10-01` 這類 `YYYY-MM-DD` 格式，不要輸入 `2026/10/01`、`10/01/2026` 或不存在的日期（例如 `2026-02-30`）。開始日期不可晚於結束日期；若沒有指定結束日期，開始日期不可晚於目前時間。修正後重新執行即可。

### PowerShell 提示禁止執行指令碼

BAT 已使用 `-ExecutionPolicy Bypass` 啟動單次 PowerShell 程序，不會永久修改系統執行原則。請只執行已確認可信的腳本；若仍遭公司群組原則或安全軟體阻擋，請洽 IT，勿自行停用安全防護。

### 無法建立 Outlook 物件，或執行時出現登入／存取提示

確認已安裝並設定傳統版 Outlook，先手動開啟並完成登入。若 Outlook 顯示程式存取警告，請依公司安全政策處理；新版 Outlook 或網頁版無法替代此腳本需要的 COM 介面。

### 找不到腳本或無法寫入檔案

確認 BAT 與 PS1 放在同一資料夾。若輸出失敗，請檢查磁碟空間、資料夾寫入權限與路徑長度；可改用較短且可寫入的 `-OutputPath`。單封郵件失敗的細節會顯示在警告訊息中。

## 資料保護

匯出檔包含完整郵件及附件，可能含個人或公司機密。請存放在經授權的位置，依組織的備份、存取與保留政策管理，不要提交到公開 Git 儲存庫或上傳到未核准的服務。

## 上傳工具到 GitHub

只將本工具所在的 `Mail` 資料夾建立為 Git 儲存庫，不要在上一層資料夾初始化，以免一併上傳其他資料。

1. 在 GitHub 建立空白儲存庫，例如 `outlook-mail-export`，不要預先加入 README、`.gitignore` 或授權檔。公司用途請先確認政策，建議設為 **Private**；私人儲存庫也不應存放郵件或機密。
2. 在 PowerShell 切換到本工具資料夾，執行以下指令。請將路徑及 `YOUR_USERNAME` 換成自己的值：

   ```powershell
   Set-Location "C:\Users\YOUR_NAME\Desktop\Test\Mail"
   git init
   git branch -M main
   git add -- .gitignore README.md Export-OutlookSender.ps1 Run-OutlookExport.bat
   git diff --cached --stat
   git diff --cached
   ```

3. 確認暫存內容只有工具及文件，沒有真實郵件、附件、密碼或公司資料，再執行：

   ```powershell
   git commit -m "Initial commit: Outlook mail export tool"
   git remote add origin https://github.com/YOUR_USERNAME/outlook-mail-export.git
   git push -u origin main
   ```

   若 Git 提示缺少身分，先設定本儲存庫的提交身分，再重試 `git commit`：

   ```powershell
   git config user.name "YOUR_NAME"
   git config user.email "YOUR_VERIFIED_OR_GITHUB_NOREPLY_EMAIL"
   ```

   推送時依 Git 的登入提示完成 GitHub 驗證，不要將密碼或權杖寫入檔案或遠端網址。

後續更新：

```powershell
git add -- .gitignore README.md Export-OutlookSender.ps1 Run-OutlookExport.bat
git diff --cached
git commit -m "Update Outlook mail export tool"
git push
```

[.gitignore](.gitignore) 已排除預設的 `outlook-export` 資料夾、常見郵件／信箱檔案、環境變數檔、日誌及編輯器暫存檔。若自訂輸出位置在儲存庫內，請另外加入對應資料夾規則，避免獨立附件或其他資料被追蹤。可用 `git status --short --ignored` 檢查排除結果。

**注意：`.gitignore` 不會停止追蹤已加入 Git 的檔案，也不會移除歷史紀錄。** 若敏感資料已被提交或推送，請先停止推送並依組織政策處理；僅新增忽略規則或刪除目前檔案不足以清除歷史中的資料。