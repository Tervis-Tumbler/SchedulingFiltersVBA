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

$SchedulingPCComputerNames = @("Scheduling2-pc")

function get-SchedulingPCComputerNames {
    $SchedulingPCComputerNames
}

function Get-SchedulingFilterButtonEvents {
    foreach ($SchedulingPCComputerName in get-SchedulingPCComputerNames) {
        $EventLogEntries = get-eventlog -LogName Application -ComputerName $SchedulingPCComputerName -Source WSH
    }

}