$SourcePath = "\\rb-sccm-pkg-d.bosch.com\rbcm$\RBCM2127\Minitab18\Temp\NMA9HC\tool\Dev\SCCM_tool.ps1"
if (Test-Path -Path "$SourcePath" -PathType Any -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
{
    $Content = Get-Content -Path $SourcePath -Raw
    Invoke-Expression -Command $Content
}
