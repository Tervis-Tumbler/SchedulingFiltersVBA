function Test-FileLock {
    [cmdletbinding()]
    param(
        [parameter(Mandatory)]
        [string]
        $Path
    )

    $oFile = New-Object System.IO.FileInfo $Path

    if ((Test-Path -Path $Path) -eq $false){
        return $false
    }

    try {
        $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        if ($oStream) {
            $oStream.Close()
        }
        $false
        }

    catch {
        # file is locked by a process.
        return $true
    }
}

function Export-SchedulingFilterExcelVBA {
    $SourceFile = "C:\Users\alozano\Documents\WindowsPowerShell\Modules\SchedulingFiltersVBA\Worksheets\SchedulingFilters.xlsm"
    $DestinationFolder = "C:\Users\alozano\Documents\WindowsPowerShell\Modules\SchedulingFiltersVBA\Worksheet Export"

    Export-ExcelProject -WorkbookPath $SourceFile -OutputPath $DestinationFolder -Verbose
}

#Export-SchedulingFilterExcelVBA

$SchedulingComputerNames = @("Scheduling2-pc","MMuniz-PC","lasbury2-pc")

function get-SchedulingComputerNames {
    $SchedulingComputerNames
}

function Get-SchedulingFilterButtonEvents {    
    foreach ($SchedulingComputerName in get-SchedulingComputerNames) {
        $EventLogEntries = get-eventlog -LogName Application -ComputerName $SchedulingComputerName -Source WSH
        $PathToStoreEvents = "\\tervis.prv\applications\Logs\Infrastructure\SchedulingFiltersVBA\$SchedulingComputerName"
        if ((Test-Path $PathToStoreEvents) -eq $False) { New-item -ItemType Directory $PathToStoreEvents }

        foreach ($EventLogEntry in $EventLogEntries) {
            $MessageProperties = $EventLogEntry.Message | ConvertFrom-Json
            [pscustomobject][ordered]@{
                MachineName = $EventLogEntry.MachineName
                FunctionName = $MessageProperties.FunctionName
                TimeGenerated = $EventLogEntry.TimeGenerated
                TimeWritten = $EventLogEntry.TimeWritten
            } | 
            ConvertTo-Json |
            Out-File -FilePath "$PathToStoreEvents\$(get-date -Format -- FileDateTime).json"

            #[datetime]::ParseExact($(get-date -Format -- FileDateTime),"yyyyMMddTHHmmssffff", [System.Globalization.CultureInfo]::CurrentCulture)
        }
    }
}