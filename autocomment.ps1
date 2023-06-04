using namespace OpenQA.Selenium

$mylocation = Get-Location

Import-Module "$mylocation\WebDriver.dll"



function cookiejsonfilecheck {
    param($filename)
    $path = Get-ChildItem -Path $mylocation -Filter $cookiefilename -Recurse | ForEach-Object{$_.FullName}
    if ($path) {
        return $path
    }
    else {
        return $false
    }
}

function setChromeDriver([string]$filename) {
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver
    $driver.Navigate().GoToUrl("$script:siteurl")
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

$ChromeDriver = setChromeDriver($cookiefilename, $siteurl)
$Actions = New-Object -TypeName Interactions.Actions ($ChromeDriver)
$ChromeDriver.Navigate().GoToUrl($onseiurl)
$items = $ChromeDriver.FindElement([By]::Id('main'))
$items = $items.FindElements([By]::ClassName(('t')))
$items = $items.FindElements([By]::ClassName(('tr3')))

$commentedpostid = @(Get-Content $mylocation\commentedpostid.txt)

$i = 0
#處理今日置頂帖
foreach ($item in $items) {
    #for top post
    try {
        $post = $item.FindElements([By]::TagName("td"))[1]
        $top = $post.FindElement([By]::TagName("img")).GetAttribute("title")
        if ($top -eq "置顶帖标志") {
            Write-Host "This is top post"
        }
        else {
            Write-Host "This is normal post"
            
        }
        $post_id = $post.GetAttribute("id")
        Write-Host $top, $post_id
        if ($post_id -in $commentedpostid) {
            Write-Host "this post is commented!"
            continue
        }
        #for today post
        try {
            $today_post = $post.FindElement([By]::CssSelector("[color='red']"))
            $post_url = $post.FindElement([By]::TagName("h3")).FindElement([By]::TagName("a")).GetAttribute('href')
            $ChromeDriver.Navigate().GoToUrl($post_url)
            # $ChromeDriver.Manage().Timeouts().ImplicitWait(5)
            $Actions.MoveToElement($ChromeDriver.FindElement([By]::Id('footer'))).Build().Perform()
            $ChromeDriver.FindElement([By]::TagName('textarea')).SendKeys($poststring)
            Start-Sleep 3
            $ChromeDriver.FindElement([By]::CssSelector(("[name='Submit']"))).Click()
            Start-Sleep 6
            $ChromeDriver.Navigate().Back()
            $ChromeDriver.Navigate().Back()
            $commentedpostid += $post_id
            Start-Sleep 3
        }
        #for other day post
        catch {
            Write-Host "other day post"
        }
        
    }
    #for not top post
    catch {
        Write-Host $Error[0]
        Write-Output "not top post or just an announcement"
    }

    $i += 1
}

$commentedpostid | Out-File $mylocation\commentpostid.txt


