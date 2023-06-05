using namespace OpenQA.Selenium

$mylocation = Get-Location

Import-Module "$mylocation\WebDriver.dll"



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
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver
    $driver.Navigate().GoToUrl($script:siteurl)
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


function isSpecialPost {
    param($post)
    #is a top post
    try {
        $post = $post.FindElements([By]::TagName("td"))[1]
        $posttype = $post.FindElement([By]::TagName("img")).GetAttribute("title")
        if ([string]::IsNullOrEmpty($posttype)) { return $false }
        if ($posttype -eq "新帖标志") {return "一般帖(新帖)"}
        else {return "置頂帖"}
        # Write-Host $post
        # Write-Host $posttype.GetType().name
    }
    #is an essence post
    catch {
        <#Do this if a terminating exception happens#>
        $script:NORMALPOSTFLAG = $true
        return $false
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

function handlePost {
    param($post, $posttype, $wantdate)
    $commentdate = handleDate($post)
    $post_id = $post.FindElements([By]::TagName("td"))[1].GetAttribute("id")
    # $today_post = $post.FindElement([By]::CssSelector("[color='red']"))
    $post_url = $post.FindElement([By]::TagName("h3")).FindElement([By]::TagName("a")).GetAttribute('href')
    $genre = $post.FindElements([By]::TagName("td"))[1].FindElement([By]::TagName("a")).Text
    $title = $post.FindElements([By]::TagName("td"))[1].FindElements([By]::TagName("a"))[1].Text
    
    $ChromeDriver.Navigate().GoToUrl($post_url)
    # $ChromeDriver.Manage().Timeouts().ImplicitWait(5)

    Write-Host "-----" -NoNewline
    if ($commentdate -ne $wantdate) { Write-Host " $commentdate" -ForegroundColor Yellow -NoNewline}
    else {Write-Host " $commentdate" -ForegroundColor Green -NoNewline}
    if ($posttype -eq "一般帖" -or $posttype -eq "置頂帖") {Write-Host " $posttype      " -NoNewline}
    else{Write-Host " $posttype" -NoNewline}
    Write-Host " $post_id" -NoNewline
    Write-Host " -----"
    Write-Host "種類: $genre"
    Write-Host "標題: $title"

    Write-Host "Posting..." -NoNewline
    # $Actions.MoveToElement($ChromeDriver.FindElement([By]::Id('footer'))).Build().Perform()
    # $ChromeDriver.FindElement([By]::TagName('textarea')).SendKeys($poststring)
    # Start-Sleep 3
    # $ChromeDriver.FindElement([By]::CssSelector(("[name='Submit']"))).Click()
    # Start-Sleep 6
    Write-Host "...done."
    $ChromeDriver.Navigate().Back()
    # $ChromeDriver.Navigate().Back()
    # Start-Sleep 3
}

$ChromeDriver = setChromeDriver($cookiefilename, $siteurl)
$ChromeDriver.Navigate().GoToUrl($onseiurl)
$actions = New-Object -TypeName Interactions.Actions ($ChromeDriver)

#確保post的elements

$i = 0
#處理今日置頂帖
$items = getPostElementList

foreach ($item in $items) {
    # Write-Host $item.GetType().name
    $posttype = isSpecialPost($item)
    if (!$posttype) { $posttype = "一般帖" }
    handlePost $item $posttype "2023-06-04"
    # $commentedpostid += $post_id
    $i += 1
}

$commentedpostid | Out-File $mylocation\commentpostid.txt


