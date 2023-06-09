# SouthPlusAutoComment
## 修改中...
## 聲明
此為個人學術性質的研究，如因為使用而導致損失，本人一概不負相關責任。

## 功能
- 自動回覆某個板上的文章 (預設為當日)
- 自行設定
    - -url 某個版面
    - -wantdate 某個日期
- Windows工作排程自動化


## 文件準備
> southpluscookie.json

將cookie存成json檔案，能存多少就存多少，目前的功能只處理name and value，其他domainame or expired time都還沒處理過。
```json
{
    "_gid" : "",
    "_ga" : "",
    "_gat" : "",

}
```
> commentedpost.txt

儲存回覆過的文章Id
```txt
td_xxxxxxx
td_xxxxxxx
td_xxxxxxx
```
> Selenium Library

從NuGet下載後把`dll`檔放到當前的工作資料夾

> ChromeWebDriver.exe

也是下載下來後把`exe`檔放到當前工作的資料夾
## Selenium 和 WebDriver 版本

Selenium `4.0.0`


## 使用方法
`autocomment.ps1`可單獨使用，請自行填寫文件內的變數
```ps1
$siteurl = ""
$poststring = ""
```
若要進行自動排程請自行填寫文件內的變數
```ps1
$tasks_urls = @{
    "任務名稱1" = "對應任務的版面網址1"
    "任務名稱2" = "對應任務的版面網址2"
}
```

## 移除
1. 若只單獨使用`autocomment.ps1`，直接刪除即可
2. 若進行了Windows工作排程，請打開工作排程器將任務資料夾右鍵刪除即可
    - 若覺得刪不乾淨，可連同機碼也一起刪除:
        ```
        HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\SouthPlus自動回覆
        ```
3. 若不知道上述如何操作，請打開`PowerShell`並執行以下命令
```ps1
Remove-Item -Path "C:\Windows\System32\Tasks\SouthPlus自動回覆"
Remove-Item -Path "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree\SouthPlus自動回覆" -Force
```