param(
  [Parameter(Mandatory=$true)]
  [string]$PptxPath,

  [Parameter(Mandatory=$true)]
  [string]$OutDir,

  [int]$Width = 1920,
  [int]$Height = 1080
)

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$powerpoint = New-Object -ComObject PowerPoint.Application
$presentation = $null

try {
  $presentation = $powerpoint.Presentations.Open($PptxPath, $true, $false, $false)
  $presentation.Export($OutDir, "JPG", $Width, $Height)
}
finally {
  if ($presentation -ne $null) {
    $presentation.Close()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null
  }
  $powerpoint.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
}

Get-ChildItem -LiteralPath $OutDir -Filter "*.JPG" | Select-Object FullName, Length, LastWriteTime
