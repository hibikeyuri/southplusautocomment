using namespace OpenQA.Selenium

param(
    [Parameter(Position=0,mandatory=$true)]
    [string] $url,
    [Parameter(Position=1,mandatory=$true)]
    [string] $wantdate
)

$mylocation = Get-Location
Write-Host $mylocation
Import-Module "$mylocation\WebDriver.dll"



$cookiefilename = "southpluscookie.json"
$commentedpostid = @(Get-Content $mylocation\commentedpostid.txt)

$NORMALPOSTFLAG = $false
$DATEFLAG = $false
$GREATERTHANFLAG = $false

$before_post_time_interval = 2
$before_submit_time_interval = 6
$wait_for_net_post_interval = 2

$siteurl = ""
$poststring = ""

$seleniumWait = ""
$opt = New-Object -TypeName Chrome.ChromeOptions
$opt.PageLoadStrategy = "eager"

function checkUrl {
    param($url)
    if ($url.Contains("fid-128")) {return "同人音声"}
    elseif ($url.Contains("fid-4")) {return "動畫資源"}
    elseif ($url.Contains("fid-14")) {return "CG資源"}
    elseif ($url.Contains("fid-6")) {return "遊戲資源"}
}

function cookiejsonfilecheck {
    param([string]$filename)
    $path = Get-ChildItem -Path $mylocation -Filter $cookiefilename -Recurse | ForEach-Object { $_.FullName }
    if ($path) {
        return $path
    }
    else {
        return $false
    }
}

function setChromeDriver {
    param([string]$filename)
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($opt)
    $driver.Navigate().GoToUrl($script:siteurl)
    Start-Sleep 2
    $script:seleniumWait = New-Object -TypeName Support.UI.WebDriverWait($driver, (New-TimeSpan -Seconds 10))
    #support for name and value, did not consider other fields like expire time, domain...etc.
    $cookiepath = cookiejsonfilecheck($filename)
    if ($cookiepath) {
        $cookiejsonobj = Get-Content $cookiepath | ConvertFrom-Json
        Write-Host $cookiejsonobj
        foreach ($cookie in $cookiejsonobj.psobject.Properties) {
            $cook = New-Object -TypeName OpenQA.Selenium.Cookie -ArgumentList $cookie.name, $cookie.value
            $driver.Manage().Cookies.AddCookie($cook)
        }
        
        return $driver
    }
    else {
        Write-Host "driver false"
        return $false
    }
}

function setChromeDriverActions {
    $Actions = New-Object -TypeName Interactions.Actions ("$script:ChromeDriver")
    return $Actions
}

#確保post的elements
function getPostElementList {
    $posts = $script:ChromeDriver.FindElement([By]::Id('main'))
    $posts = $posts.FindElements([By]::ClassName(('t')))
    $posts = $posts.FindElements([By]::ClassName(('tr3')))
    return $posts
}

function whatSpecialPost {
    param($post)
    #is a top post
    try {
        $post = $post.FindElements([By]::TagName("td"))[1]
        $posttype = $post.FindElement([By]::TagName("img")).GetAttribute("title")
        if ([string]::IsNullOrEmpty($posttype)) { 
            $script:NORMALPOSTFLAG = $true
            return "一般帖" 
        }
        if ($posttype -eq "新帖标志") {
            $script:NORMALPOSTFLAG = $true
            return "一般帖(新帖)"
        }
        else {
            $script:NORMALPOSTFLAG = $false
            return "置頂帖"
        }
    }
    #is an essence post or normal post
    catch {
        $script:NORMALPOSTFLAG = $true
        return "一般帖"
    }
}

function handleDate {
    param($post)
    #is a normal post
    try {
        $date = $post.FindElements([By]::TagName('td'))[2].FindElement([By]::TagName('div')).Text
        $script:DATEFLAG = $true
        return $date
    }
    #is an announcement post
    catch {
        return $false
    }
}

function postInfo {
    param($post, $posttype, $wantdate)

    $commentdate = handleDate($post)
    $post_id = $post.FindElements([By]::TagName("td"))[1].GetAttribute("id")
    # $today_post = $post.FindElement([By]::CssSelector("[color='red']"))
    try {
        $post_url = $post.FindElement([By]::TagName("h3")).FindElement([By]::TagName("a")).GetAttribute('href')
        $genre = $post.FindElements([By]::TagName("td"))[1].FindElement([By]::TagName("a")).Text
        $title = $post.FindElements([By]::TagName("td"))[1].FindElements([By]::TagName("a"))[1].Text
        $samedate = $($(Get-Date -Format "yyyy-MM-dd") -eq $commentdate)
    }
    catch {
        Write-Host "讀取錯誤，可能是公告類帖子"
    }
    # $ChromeDriver.Manage().Timeouts().ImplicitWait(5)

    Write-Host "-----" -NoNewline
    if (!([string]::IsNullOrEmpty($wantdate))) {
        if ($commentdate -eq $wantdate) {
            Write-Host " $commentdate" -ForegroundColor Green -NoNewline
        }
        elseif ($($(Get-Date -Format "yyyy-MM-dd") -eq $commentdate)) {
            Write-Host " $commentdate" -ForegroundColor Cyan -NoNewline
        } 
        else {
            Write-Host " $commentdate" -ForegroundColor Yellow -NoNewline
        }
    }
    else {
        if ($samedate){
            Write-Host " $commentdate" -ForegroundColor Green -NoNewline
        }
        else {
            Write-Host " $commentdate" -ForegroundColor Yellow -NoNewline
        }
    }
    if ($posttype -eq "一般帖" -or $posttype -eq "置頂帖") {Write-Host " $posttype      " -NoNewline}
    else{Write-Host " $posttype" -NoNewline}
    Write-Host " $post_id" -NoNewline
    Write-Host " -----"
    Write-Host "種類: $genre"
    Write-Host "標題: $title"

    #回傳張貼日期、帖子id、帖子URL、帖子種類、帖子標題
    return $commentdate, $post_id, $post_url, $genre, $title
}

function handlePost {
    Write-Host "Posting..." -NoNewline
    $script:Actions.MoveToElement($script:ChromeDriver.FindElement([By]::Id('footer'))).Build().Perform()
    $script:ChromeDriver.FindElement([By]::TagName('textarea')).SendKeys($script:poststring)
    Start-Sleep $script:before_post_time_interval
    $script:ChromeDriver.FindElement([By]::CssSelector(("[name='Submit']"))).Click()
    Start-Sleep $script:before_submit_time_interval
    Write-Host "...done."
    $script:ChromeDriver.Navigate().Back()
    $script:ChromeDriver.Navigate().Back()
    Start-Sleep $script:wait_for_net_post_interval
}

function checkPost {
    param($post, $posttype, $wantdate)
    
    $commentdate, $post_id, $post_url, $genre, $title = postInfo $post $posttype $wantdate

    #有給日期
    if (![string]::IsNullOrEmpty($wantdate)) {
        if ($wantdate -gt $commentdate) { $script:GREATERTHANFLAG = $true}
        else { $script:GREATERTHANFLAG = $false}
        if ($wantdate -eq $commentdate -and $post_id -notin $script:commentedpostid) {
            #$script:ChromeDriver.Manage().Timeouts().ImplicitWait = 5
            $script:ChromeDriver.Navigate().GoToUrl($post_url)
            handlePost
            $script:commentedpostid += $post_id
            $commentedpostid | Out-File $mylocation\commentedpostid.txt
            return $true
        }
        else {
            return $false
        }
    }
    #沒給日期
    else {
        if ($(Get-Date -Format "yyyy-MM-dd") -gt $commentdate) { $script:GREATERTHANFLAG= $true}
        else { $script:GREATERTHANFLAG = $false}
        if ($($(Get-Date -Format "yyyy-MM-dd") -eq $commentdate) -and $post_id -notin $script:commentedpostid) {
            $script:ChromeDriver.Navigate().GoToUrl($post_url)
            handlePost
            $script:commentedpostid += $post_id
            $commentedpostid | Out-File $mylocation\commentedpostid.txt
            return $true
        }
        else {
            return $false
        }
    }
    
}

$ChromeDriver = setChromeDriver($cookiefilename, $siteurl)
$Actions = New-Object -TypeName Interactions.Actions ($ChromeDriver)


$pagedurl = $url.Substring(0, $url.LastIndexOf('.'))
$pagedurl += "-page"
$i = 1
:outer while ($true) {
    # $wantdate = "2023-06-05"
    #fid128-page-1.html
    $pagedurl += "-$i.html"
    Write-Host $pagedurl
    $postbase = checkUrl($url)
    Write-Host "-------------------" -ForegroundColor Cyan
    Write-Host "正在進行進行 $postbase 板第 $i 頁..." -ForegroundColor Cyan
    Write-Host "-------------------" -ForegroundColor Cyan
    $ChromeDriver.Navigate().GoToUrl($pagedurl)
    $items = getPostElementList
    foreach ($item in $items) {
    # Write-Host $item.GetType().name
        $posttype = whatSpecialPost($item)
        $check = checkPost $item $posttype $wantdate

        Write-Host "一般帖: $NORMALPOSTFLAG, 有日期: $DATEFLAG, 設定日期大於帖子張貼日期: $GREATERTHANFLAG"

        if ($($posttype -eq "一般帖(新帖)" -or $posttype -eq "一般帖") -and $NORMALPOSTFLAG -and $DATEFLAG -and $GREATERTHANFLAG) {
            Write-Host "=====      Auto comment is done.      =====" -ForegroundColor Cyan
            break outer
        }

        if ($check) {
            Write-Host "-----             回覆已完成             -----" -ForegroundColor Green
        }
        else {
            Write-Host "-----       非自選日期或是已回覆        -----" -ForegroundColor Red
        }
        Write-Host "`r`n"
    } 

    $pagedurl = $pagedurl.Substring(0, $pagedurl.LastIndexOf('-'))
    Write-Host $pagedurl
    $i += 1
}

$ChromeDriver.Close()
