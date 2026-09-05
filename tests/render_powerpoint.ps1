param(
    [Parameter(Mandatory=$true)][string]$PptxPath,
    [Parameter(Mandatory=$true)][string]$OutputDirectory
)
$ErrorActionPreference = 'Stop'
$inputFile = (Resolve-Path -LiteralPath $PptxPath).Path
$outputFolder = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
$existingPowerPoint = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
$powerPoint = $null
$presentation = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $presentation = $powerPoint.Presentations.Open($inputFile, -1, 0, 0)
    $slideCount = $presentation.Slides.Count
    for ($slideIndex = 1; $slideIndex -le $slideCount; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        $slide.Export((Join-Path $outputFolder "slide-$slideIndex.png"), 'PNG', 1600, 900)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($slide)
    }
    Write-Output "PowerPoint opened and rendered $slideCount slides."
} finally {
    if ($null -ne $presentation) {
        $presentation.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    if ($null -ne $powerPoint) {
        if (-not $existingPowerPoint) { $powerPoint.Quit() }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
    }
}
