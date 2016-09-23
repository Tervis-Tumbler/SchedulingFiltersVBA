function Export-SchedulingFilterExcelVBA{
    $SourceFile = "C:\Users\alozano\Documents\WindowsPowerShell\Modules\SchedulingFiltersVBA\Worksheets\SchedulingFilters.xlsm"
    $DestinationFolder = "C:\Users\alozano\Documents\WindowsPowerShell\Modules\SchedulingFiltersVBA\Worksheet Export"

    Export-ExcelProject -WorkbookPath $SourceFile -OutputPath $DestinationFolder -Verbose
}