<#

BackupNewspage.ps1

    2018-08-12 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = (Get-WJEmail -Name lyu).MailAddress
$mailSbj    = $scriptName
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 50 -Path $log

###################################################################################





$newspage = (Get-WJPath -Name newspage).Path
$workDate = (Get-Date).AddDays(-14)
$workPath = ($newspage + $workDate.ToString("yyyyMMdd"))
$bBanPath = ($newspage + "¸Éª©\" + $workDate.ToString("yyyyMMdd"))
$bBanIncl = @("4259*.pdf", "4257*.pdf", "4267*.pdf")
$bBanExcl = @("*wk.pdf")

Write-Log -Verb "newspage" -Noun $newspage -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal
Write-Log -Verb "workPath" -Noun $workPath -Path $log -Type Short -Status Normal
Write-Log -Verb "bBanPath" -Noun $bBanPath -Path $log -Type Short -Status Normal
Write-Log -Verb "bBanIncl" -Noun ($bBanIncl -join ", ") -Path $log -Type Short -Status Normal
Write-Log -Verb "bBanExcl" -Noun ($bBanExcl -join ", ") -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log



# Find all 4259* 4257* 4267* pdf files

if(Test-Path $workPath){
    $items = @(Get-ChildItem -Path $workPath -Include $bBanIncl -Exclude $bBanExcl -Recurse)
}else{
    $items = @()
    $mailMsg = $mailMsg + (Write-Log -Verb "NO FOLDER" -Noun $workPath -Path $log -Type Long -Status Bad -Output String) + "`n"
    $hasError = $true
}



if($items.Count -ne 0){

    if(!(Test-Path $bBanPath)){
        New-Item $bBanPath -ItemType Directory | Out-Null
        Write-Log -Verb "CREATE" -Noun $bBanPath -Path $log -Type Long -Status Normal
    }

    foreach($i in $items){

        $copyFrom = $i.FullName
        $copyTo = $bBanPath + "\" + $i.Name
        Write-Log -Verb "copyFrom" -Noun $copyFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "copyTo" -Noun $copyTo -Path $log -Type Short -Status Normal

        try{

            Write-Log -Verb "COPY FROM" -Noun $copyFrom -Path $log -Type Long -Status Normal
            Copy-Item $copyFrom $copyTo -ErrorAction Stop
            Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Good

        }catch{

            $mailMsg = $mailMsg + (Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Bad -Output String) + "`n"
            $hasError = $true

        }

    }

}else{

    Write-Log -Verb "NO BBAN FILE" -Noun $workPath -Path $log -Type Long -Status Normal

}





###################################################################################

# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $bBanPath = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $bBanPath -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $bBanPath -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam