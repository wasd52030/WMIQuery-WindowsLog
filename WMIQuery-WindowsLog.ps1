function Check-Administrator {
    $context = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $context.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $scriptPath = $MyInvocation.ScriptName
        $process = New-Object System.Diagnostics.ProcessStartInfo
        $process.FileName = "powershell.exe"
        $process.Arguments = "-File `"$scriptPath`" -type $type"
        $process.Verb = "runas"
        [System.Diagnostics.Process]::Start($process) | Out-Null
        Exit
    }
}


Check-Administrator

$types = @('Application', 'Security' , 'System')
$logNames = @{ Application = "應用程式"; Security = "安全性"; System = "系統" }

# reference -> https://learn.microsoft.com/en-us/windows/win32/wmisdk/cim-datetime?redirectedfrom=MSDN
# GMT+8 -> 60*8 = 480
$yesterday = (Get-Date).AddDays(-1)
$start = $yesterday.Date.ToString("yyyyMMddHHmmss")+".000000+480"
$end = $yesterday.Date.AddHours(23).AddMinutes(59).AddSeconds(59).ToString("yyyyMMddHHmmss")+".000000+480"
$kmtDate = "$($yesterday.Year-1911)$($yesterday.Date.ToString("MMdd"))"


for($i=0;$i -lt $types.Length;$i++){
    # Write-Host "SELECT * FROM Win32_NTLogEvent WHERE Logfile = '$($types[0])' And TimeGenerated > '$start' And TimeGenerated < '$end'"
    $log = Get-CimInstance -query "SELECT * FROM Win32_NTLogEvent WHERE Logfile = '$($types[$i])' And TimeGenerated > '$start' And TimeGenerated < '$end'" `
            | Select-Object -Property Type, `
                                      TimeWritten, `
                                      SourceName, `
                                      EventCode, `
                                      @{Name="CategoryString"; Expression = { if ($_.CategoryString) { $_.CategoryString } else { "無" } }}, 
                                      Message `
            | ConvertTo-Csv -NoTypeInformation `
            | Select-Object -Skip 1 `
            | ForEach-Object { $_ -replace '"([^"]*)"', '$1' }

    $FilePath="~\Downloads\$($kmtDate)$($logNames[$types[$i]]).csv"

    if($types[$i] -eq "Security"){
        Set-Content -Path $FilePath -Value "關鍵字,日期和時間,來源,事件識別碼,工作類別"
    }else{
        Set-Content -Path $FilePath -Value "等級,日期和時間,來源,事件識別碼,工作類別"
    }

    Add-Content -Path $FilePath -Value $log

    Write-Host "已取得$($logNames[$types[$i]])紀錄！"
}


pause