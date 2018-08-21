<#

BackupNewspage.ps1

    2018-08-12 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Get-Module -ListAvailable WorldJournal.* | Remove-Module -Force
Get-Module -ListAvailable WorldJournal.* | Import-Module -Force

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
$n_backup = (Get-WJPath -Name newspage_backup).Path
$workDate = (Get-Date).AddDays(-15)
$workPath = ($newspage + $workDate.ToString("yyyyMMdd"))
$bkupPath = ($n_backup + $workDate.ToString("yyyyMMdd"))
$bBanPath = ($newspage + "¸Éª©\" + $workDate.ToString("yyyyMMdd"))
$bBanIncl = @("4259*.pdf", "4257*.pdf", "4267*.pdf")
$bBanExcl = @("*wk.pdf")

Write-Log -Verb "newspage" -Noun $newspage -Path $log -Type Short -Status Normal
Write-Log -Verb "n_backup" -Noun $n_backup -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal
Write-Log -Verb "workPath" -Noun $workPath -Path $log -Type Short -Status Normal
Write-Log -Verb "bkupPath" -Noun $bkupPath -Path $log -Type Short -Status Normal
Write-Log -Verb "bBanPath" -Noun $bBanPath -Path $log -Type Short -Status Normal
Write-Log -Verb "bBanIncl" -Noun ($bBanIncl -join ", ") -Path $log -Type Short -Status Normal
Write-Log -Verb "bBanExcl" -Noun ($bBanExcl -join ", ") -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log



# Find all $bBanIncl pdf files

if(Test-Path $workPath){

    $items = @(Get-ChildItem -Path $workPath -Include $bBanIncl -Exclude $bBanExcl -Recurse)

}else{

    $items = @()
    $mailMsg = $mailMsg + (Write-Log -Verb "NO FOLDER" -Noun $workPath -Path $log -Type Long -Status Bad -Output String) + "`n"
    $hasError = $true

}



# Copy $bBanIncl pdf files

if($items.Count -ne 0){

    if(!(Test-Path $bBanPath)){

        New-Item $bBanPath -ItemType Directory | Out-Null
        Write-Log -Verb "NEW" -Noun $bBanPath -Path $log -Type Long -Status Normal

    }

    $items | Copy-Files -From $workPath -To $bBanPath | ForEach-Object{
        Write-Log -Verb "copyFrom" -Noun $_.CopyFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "copyTo" -Noun $_.CopyTo -Path $log -Type Short -Status Normal
        Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status
    }

    <#
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
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
            $hasError = $true

        }

    }
    #>

}else{

    Write-Log -Verb "NO BBAN FILE" -Noun $workPath -Path $log -Type Long -Status Normal

}



Write-Line -Length 50 -Path $log



# Create $bkupPath

if(!(Test-Path $bkupPath)){

    New-Item $bkupPath -ItemType Directory | Out-Null
    Write-Log -Verb "NEW" -Noun $bkupPath -Path $log -Type Long -Status Normal

}



# Move-Files from $workPath to $bkupPath

Get-ChildItemPlus $workPath | Sort-Object -Descending | Move-Files -From $workPath -To $bkupPath | ForEach-Object{

    Write-Log -Verb "moveFrom" -Noun $_.MoveFrom -Path $log -Type Short -Status Normal
    Write-Log -Verb "moveTo" -Noun $_.MoveTo -Path $log -Type Short -Status Normal
    Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status

    if($_.Status -eq "Bad"){

        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status -Output String) + "`n"
        $hasError = $true

    }

}



# Delete $workPath

if($hasError){

    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE SKIP" -Noun $workPath -Path $log -Type Long -Status Bad -Output String) + "`n"

}else{

    Write-Log -Verb "REMOVE" -Noun $workPath -Path $log -Type Long -Status Normal

    try{

        $temp = $workPath
        Remove-Item $workPath -Recurse -Force -ErrorAction Stop
        Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good

    }catch{

        $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
        $hasError = $true

    }

}



###################################################################################

# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
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