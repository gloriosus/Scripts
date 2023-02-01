$RootFolder = (Split-Path -Path $PSScriptRoot -Parent)
$Config = Get-Content -Path "$($RootFolder)\Config\Publish-122Data.json" | ConvertFrom-Json

$WebDriverPath = $Config.WebDriverPath
$LogPath = $Config.LogPath
$Username = $Config.Username
$Password = $Config.Password
$SmtpServer = $Config.SmtpServer
$SmtpPort = $Config.SmtpPort
[string]$SmtpUser = $Config.SmtpUser
[securestring]$SmtpPassword = ConvertTo-SecureString $Config.SmtpPassword -AsPlainText -Force
$SenderEmail = $Config.SenderEmail
$MailList = $Config.MailList

$DayOfWeek = (Get-Date).DayOfWeek

if (($DayOfWeek -eq "Sunday") -or ($DayOfWeek -eq "Monday")) {
    $AllOperator = 0
    $DayOperator = 0
} else {
    $AllOperator = 2
    $DayOperator = 2
}

try {
    Add-Type -Path "$($WebDriverPath)\WebDriver.dll"

    $YandexDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver -ArgumentList $WebDriverPath
    $YandexDriver.Navigate().GoToUrl("https://scmks.ru/treatment/table")

    Start-Sleep -Seconds 5
    # Log in
    $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("//*[@id=`"Username`"]")).SendKeys($Username)
    $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("//*[@id=`"Password`"]")).SendKeys($Password)
    $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[3]/div/div[2]/div/div/div/form/fieldset/div[3]/div/button")).Click()

    Start-Sleep -Seconds 5
    # Click Create button
    $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("//*[@id=`"app`"]/div/div[2]/main/div/div[2]/div/div/div[1]/div/div[4]/button")).Click()

    Start-Sleep -Seconds 5
    # Set date and select System Type from dropdown menu
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Отчетная дата`"]")).Click()

    if ((Get-Date).Day -eq "1") {
        $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[3]/div/div/div/div[2]/div[1]/div/div[1]/div[1]/button")).Click()
        Start-Sleep -Seconds 2
    }

    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[class=`"q-btn q-btn-item non-selectable no-outline q-btn--flat q-btn--rectangle q-btn--actionable q-focusable q-hoverable q-btn--dense`"]")).Click()
    
    $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[1]/div/div[2]/main/div/form/div/div[2]/div[2]/div[2]/label/div/div[1]/div[2]/div[1]")).Click()
    Start-Sleep -Seconds 3
    $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[3]/div/div[2]/div[3]/div[2]")).Click()

    # Fill operator data
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Кол-во операторов в смену`"]")).SendKeys($AllOperator)
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Операторов днем`"]")).SendKeys($DayOperator)
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Операторов ночью`"]")).SendKeys("0")

    # Fill call data
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Количество звонков`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Неотвеченные вызовы`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Количество обратных звонков`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Среднее время разговора`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Максимальное время ожидания`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Среднее время ожидания`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Ожидание более 3 мин, %`"]")).SendKeys("0")

    # Fill stat data
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Переключение на врача, %`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Записи к врачу, %`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Запись на анализы, %`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Запись на вакцинацию, %`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Вызов врача, %`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Вызов скорой помощи, %`"]")).SendKeys("0")
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Вызов волонтера, %`"]")).SendKeys("0")

    # Mark necessary checkboxes
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Информирование граждан по вопросам организации мед. помощи при COVID-19`"]")).Click()
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Организация вызова врача на дом`"]")).Click()
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Запись на прием к врачу`"]")).Click()
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Информирование об адресах мед. организаций для проведения лабораторных исследований и КТ`"]")).Click()
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Информирование о возможностях подачи жалоб на отказ в медиц. Помощи и качество предоставленных мед. услуг`"]")).Click()
    $YandexDriver.FindElement([OpenQA.Selenium.By]::CssSelector("[aria-label=`"Информирование о вакцинации против COVID-19`"]")).Click()

    # Submit the form
    $YandexDriver.FindElement([OpenQA.Selenium.By]::XPath("/html/body/div[1]/div/div[2]/main/div/form/div/div[6]/button[3]")).Click()

    Start-Sleep -Seconds 5
    $YandexDriver.Navigate().GoToUrl("https://scmks.ru/treatment/table")

    Start-Sleep -Seconds 5
    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd")
    New-Item -Path "$($LogPath)\$($CurrentDate)" -ItemType Directory -Force
    $Screenshot = ([OpenQA.Selenium.ITakesScreenshot]$YandexDriver).GetScreenshot()
    $Screenshot.SaveAsFile("$($LogPath)\$($CurrentDate)\result.jpg", [OpenQA.Selenium.ScreenshotImageFormat]::Jpeg)

    New-Item -Path "$($LogPath)\$($CurrentDate)\exception.log" -Force

    $YandexDriver.Close()
    $YandexDriver.Quit()

    $Result = "Успех"
} catch [System.Exception] {
    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd")

    New-Item -Path "$($LogPath)\$($CurrentDate)" -ItemType Directory -Force
    Write-Output $_.Exception.Message | Out-File -FilePath "$($LogPath)\$($CurrentDate)\exception.log" -Encoding UTF8

    New-Item -Path "$($LogPath)\$($CurrentDate)\result.jpg" -Force

    $Result = "Провал"
} finally {
    $CurrentDate = (Get-Date).ToString("yyyy-MM-dd")
    [PSCredential]$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $SmtpUser, $SmtpPassword
    Send-MailMessage -From "Отчет 122 <$($SenderEmail)>" -To $MailList -Subject "Операция 122: $($Result)" -Body "Операция завершена $(Get-Date)." -Attachments "$($LogPath)\$($CurrentDate)\result.jpg", "$($LogPath)\$($CurrentDate)\exception.log" -Encoding UTF8 -SmtpServer $SmtpServer -Port $SmtpPort -Credential $Credential
}