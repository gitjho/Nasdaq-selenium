$bDebug = $false
#$bDebug = $true
Start-Transcript -Path c:\automation\psl\nasdaq-transcript.log
$scriptpath = split-path -parent $MyInvocation.MyCommand.Definition

$YourURL = "http://www.nasdaqomxnordic.com/bonds/denmark/microsite?Instrument=XCSE0%3A0RDSD21S38" # Website we'll log to
[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')

Function UpdateChromeDriver() {
    #detect installed version of chrome
    $GCVersionInfo = (Get-Item (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe').'(Default)').VersionInfo
    $GCInstalledMajorVersion = $GCVersionInfo.ProductMajorPart
    Write-Host "Major installed Chrome version: $GCInstalledMajorVersion"


    #detect last downloaded version of chromedriver.exe
    If (Test-Path -Path "C:\Automation\PSL\chromedriver.version.txt") {
        $GCCurrentDriverVersion = Get-Content -Path "C:\Automation\PSL\chromedriver.version.txt"
    }
    Else {
        $GCCurrentDriverVersion = "0"
    }
    Write-Host "Current driver version: $GCCurrentDriverVersion"


    #detect latest release for the installed major version
    $response = Invoke-WebRequest -Uri "https://googlechromelabs.github.io/chrome-for-testing/latest-versions-per-milestone-with-downloads.json"
    If ($response.StatusCode -eq 200) { 
        $response_json = $response.content | ConvertFrom-Json
        $GCReleasedLatestVersion = $response_json.milestones.$($GCInstalledMajorVersion).version
        $GCDownloadURL = ($response_json.milestones.$($GCInstalledMajorVersion).downloads.chromedriver | Where-Object -Property platform -eq 'win32').url
    }
    Else {
        $GCReleasedLatestVersion = $null
    }
    Write-Host "Latest released version: $GCReleasedLatestVersion"

    #download update
    If ($GCReleasedLatestVersion -ine $GCCurrentDriverVersion) {
        Write-Host "Downloading chromedriver update"
        Invoke-WebRequest -Uri $GCDownloadURL -OutFile "$scriptpath\chromedriver_win32.zip"
        Write-Host "Extracting chromedriver update"
        Expand-Archive -Path "$scriptpath\chromedriver_win32.zip" -DestinationPath $scriptpath -Force
        Copy-Item -Path .\chromedriver-win32\chromedriver.exe -Destination $scriptpath -Force
        $GCReleasedLatestVersion | Out-File -FilePath "C:\Automation\PSL\chromedriver.version.txt"
        Remove-Item -Path "$scriptpath\chromedriver_win32.zip" -Force
        Remove-Item -Path "$scriptpath\chromedriver-win32\" -Force -Recurse
    }

}

UpdateChromeDriver
Add-Type -Path "$scriptpath\WebDriver.dll" # Adding Selenium's .NET assembly (dll) to access it's classes in this PowerShell session
Add-Type @"
    using System;
    using System.Runtime.InteropServices;
  
    public class shell32  {
        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(
             [MarshalAs(UnmanagedType.LPStruct)] 
             Guid       rfid,
             uint       dwFlags,
             IntPtr     hToken,
             out IntPtr pszPath
         );

         public static string GetKnownFolderPath(Guid rfid)  {
            IntPtr pszPath;
            if (SHGetKnownFolderPath(rfid, 0, IntPtr.Zero, out pszPath) != 0) {
                return "Could not get folder";
            }
            string path = Marshal.PtrToStringUni(pszPath);
            Marshal.FreeCoTaskMem(pszPath);
            return path;
         }
    }
"@

[OpenQA.Selenium.Chrome.ChromeOptions]$options = New-Object OpenQA.Selenium.Chrome.ChromeOptions

$options.AddArgument('--window-position=200,150')
$options.AddArgument('window-size=1300,900')

$ChromeDriver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($options) # Creates an instance of this class to control Selenium and stores it in an easy to handle variable
$ChromeDriver.Navigate().GoToURL($YourURL) # Browse to the specified website
Start-Sleep -Seconds 2
$ChromeDriver.FindElement([OpenQA.Selenium.By]::Id('ui-id-4')).click()
$a = $ChromeDriver.FindElement([OpenQA.Selenium.By]::XPath('/html/body/section/div/div/div/section/div/section/div/div/div[3]/div/div[1]/div/div/div'))

$Latest = $a.FindElement([OpenQA.Selenium.By]::ClassName("db-a-lsp"))
$Percent = $a.FindElement([OpenQA.Selenium.By]::ClassName("db-a-chp"))
$High = $a.FindElement([OpenQA.Selenium.By]::ClassName("db-a-hp"))
$Low = $a.FindElement([OpenQA.Selenium.By]::ClassName("db-a-lp"))
$Ref = $a.FindElement([OpenQA.Selenium.By]::ClassName("db-a-apatd"))

$params = @{
    "Latest"="$($Latest.text)";
    "Percent"="$($Percent.text)";
    "High"="$($High.text)";
    "Low"="$($Low.text)";
    "Ref"="$($Ref.text)";
    "Url"="$($YourURL)";
}
write-host "Latest price: $($Latest.text)"

If ($params.Latest -ne "") { #we have values for today. String is empy if no trades today
    $params | Export-Clixml nasdaqnew.xml
    $newxml = Import-Clixml nasdaqnew.xml

    if (-not (test-path nasdaqsaved.xml)) { @{"asdf"="fdsa";} | Export-Clixml nasdaqsaved.xml} #create dummy file if not exist

    $savedxml = Import-Clixml nasdaqsaved.xml

    if ( -not (Compare-Object $newxml.values $savedxml.values) ) {
        # object properties and values match
    }
    else {
        #object properties and values do not match. New nasdaq trade detected
        $params | Export-Clixml nasdaqsaved.xml

        $params.Add("Date", "$($(get-date).tostring())") #add date AFTER saving so they can be compared
        Invoke-WebRequest -Uri http://10.0.1.140:1880/nasdaq -Method POST -Body $params #publish new values to node-red webservice
    }
}

$ChromeDriver.Dispose()
Stop-Transcript