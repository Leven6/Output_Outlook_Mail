# Outlook 郵件匯出工具

將指定寄件者的 Outlook 郵件匯出成 **HTML**，使用瀏覽器即可閱讀。每封郵件各有一個資料夾，保存 HTML、Outlook 匯出的內文圖片資源，以及獨立的附件檔案。工具只讀取及匯出郵件，不會移動或刪除 Outlook 裡的原始郵件，也不需要安裝 Python。

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
6. 看到 `Done.` 後，查看 BAT 同層的 `outlook-export` 資料夾，進入寄件者與郵件資料夾，雙擊 `index.html` 用瀏覽器閱讀。視窗會暫停，按任意鍵關閉。

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
      20261027-093000 - Meeting notes - A1B2C3D4E5/
        index.html
        index_files/
          image001.png
        attachments/
          001 - report.pdf
          002 - screenshot.png
```

`index_files` 是 Outlook 視需要產生的圖片與其他資源資料夾，實際名稱及檔案依郵件而異；沒有附件時，不會建立 `attachments`。

不同寄件者會存入各自的資料夾；同一寄件者在不同 Outlook 子資料夾中的郵件，會集中輸出，不會重建原本的 Outlook 資料夾階層。**移動、備份或分享時，請保留整封郵件的資料夾，不要只搬走 HTML 檔，否則圖片可能無法顯示。**

### 從舊版 MSG 匯出切換

新版會重新從 Outlook 讀取郵件並匯出 HTML，**不會將磁碟上已有的 `.msg` 自動轉檔**，也不會刪除舊檔。增量判斷只讀取已完成的 HTML 郵件資料夾，不使用舊 `.msg` 的時間；若尚無 HTML 匯出且日期皆留白，會重新匯出掃描範圍內的所有符合郵件。

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

`-WhatIf` 不會儲存 HTML、圖片或附件，但仍會建立輸出根目錄及寄件者資料夾，並連線 Outlook 掃描郵件。預覽會使用指定的日期區間；未指定日期時，則套用下方的增量匯出規則。

### 參數

| 參數 | 用途 | 預設值 |
| --- | --- | --- |
| `-Sender` | 必填，寄件者的完整電子郵件地址；不分大小寫精確比對，不支援顯示名稱、部分地址或萬用字元。 | 無 |
| `-OutputPath` | 輸出根目錄，工具會在其下建立寄件者資料夾。含空白的路徑請加引號。 | PowerShell 目前工作目錄下的 `outlook-export` |
| `-Recurse` | 同時掃描收件匣的所有子資料夾。 | 不包含子資料夾；BAT 會自動加上此參數 |
| `-StartDate` | 開始日期，格式 `yyyy-MM-dd`，包含當日零點起的郵件。 | 單獨省略時無下限；起訖皆省略時使用增量模式 |
| `-EndDate` | 結束日期，格式 `yyyy-MM-dd`，包含當天整日。 | 本次啟動時間 |
| `-WhatIf` | 預覽匯出動作，不儲存 HTML、圖片或附件。 | 關閉 |
| `-Verbose` | 顯示掃描資料夾及匯出檔案的詳細訊息。 | 關閉 |

BAT 固定輸出到 BAT 所在資料夾下的 `outlook-export`；直接執行 PowerShell 腳本時，預設位置則取決於目前工作目錄。

## 匯出範圍與增量規則

- **資料夾範圍：**只掃描 Outlook 目前設定檔的預設收件匣；加上 `-Recurse` 才包含其子資料夾。不會自動掃描其他帳號、共用信箱、封存資料夾、寄件備份或整個信箱。
- **時間依據：**使用郵件的寄件時間 `SentOn`，不是收件時間；日期依電腦本地時間解讀，不接受時分秒或時區格式。未指定結束日期時，上限是本次啟動匯出時的電腦時間。
- **手動日期區間：**只要填寫任一日期，就忽略現有 HTML 匯出的最新時間，完整檢查指定區間。同一郵件資料夾內已存在 `index.html` 時會跳過，不覆寫。
- **增量模式第一次執行：**起訖日期皆留白，且寄件者輸出資料夾內沒有可辨識、內含 `index.html` 的郵件資料夾時，匯出掃描範圍內、截至啟動時間的所有符合郵件。
- **增量模式後續執行：**起訖日期皆留白時，讀取已完成的 HTML 郵件資料夾名稱中的最新時間，只處理該時間起至本次啟動時間的郵件。時間相同的郵件仍會檢查，已有 HTML 則跳過，不覆寫。
- **資料夾名稱：**`yyyyMMdd-HHmmss - 郵件主旨 - 識別碼`，HTML 固定為其中的 `index.html`。識別碼由 Outlook 的 `EntryID` 計算；主旨中的非法檔名字元會替換成 `_`，主旨最多保留 100 個字元，空白主旨使用 `no-subject`。
- **圖片處理：**使用 Outlook 原生 HTML 儲存功能，保留其匯出的內嵌圖片與相對路徑資源。僅有網路連結、未包含在郵件中的圖片，不保證可離線顯示，也不會另外主動下載。
- **附件處理：**將 Outlook 附件集合中的檔案存入該封郵件的 `attachments` 資料夾，檔名前加上 `001 -`、`002 -` 等編號，避免同名附件互相覆蓋；非法檔名字元會替換成 `_`。部分內嵌圖片本身也是附件，因此可能同時出現在圖片資源資料夾及 `attachments` 中。
- **完成判斷：**先在 `.outlook-html-...` 暫存資料夾中儲存附件及 HTML，全部成功後才改成正式郵件資料夾。單封郵件失敗時會清理該次暫存內容並繼續下一封；若程序被強制中止，可能留下暫存資料夾，但不會將它算入增量進度。

**注意：增量匯出不是完整的缺檔補齊。** 若舊郵件晚同步進來、先前匯出失敗、刪除了較舊的匯出檔，或後來才開啟子資料夾掃描，只要郵件寄件時間早於現有最新匯出時間，就不會被補匯。

需要補匯舊郵件時，可直接指定包含那些郵件的日期區間；既有 HTML 不會覆寫。需要所有郵件重新輸出成新檔時，請用 `-OutputPath` 指向新的空白輸出根目錄，並省略日期或指定涵蓋全部郵件的區間。不要混用不同信箱的匯出結果，也不要手動改動用來判斷增量時間的郵件資料夾名稱前綴。

## 完成訊息

```text
Done. Exported: 12; already present: 3; failed: 0.
Output: ...\outlook-export\sender@example.com
```

- `Exported`：本次成功匯出的郵件數。
- `already present`：本次處理時間範圍內，因同一郵件資料夾已有 `index.html` 而跳過的數量；不是所有歷次匯出的總數。
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

### HTML 開啟後圖片不見了

確認 `index.html` 與 Outlook 產生的資源資料夾仍在同一封郵件資料夾內，且沒有改名或只複製 HTML。若圖片原本依賴網路連結，離線時仍可能無法顯示；遠端圖片也可能被瀏覽器或公司的安全政策阻擋。

### 郵件資料夾存在，但 HTML 或附件已被刪除

若 `index.html` 仍在，工具會把該封視為已匯出，不會檢查圖片或附件是否齊全。請先將該封郵件資料夾移到輸出位置之外備份，再指定日期區間重新匯出。若資料夾存在但缺少 `index.html`，匯出會因目標位置已存在而失敗，不會覆蓋剩餘檔案；也請先移開備份後重試。

## 資料保護

HTML 方便閱讀，但不等同完整的 MSG 郵件封存；需要完整 Outlook 郵件物件或法遵封存時，請另外保留原始 MSG 或使用核准的封存系統。匯出的內容及附件可能含個人或公司機密，請存放在經授權的位置，依組織的備份、存取與保留政策管理，不要提交到公開 Git 儲存庫或上傳到未核准的服務。開啟 HTML 可能載入遠端圖片，也請勿執行不明附件。