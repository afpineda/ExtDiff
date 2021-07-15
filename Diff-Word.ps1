param(
    [Parameter(Mandatory=$true)]
    [string] $BaseFileName,
    [Parameter(Mandatory=$true)]
    [string] $ChangedFileName
)

$ErrorActionPreference = 'Stop'

# Constants
$wdDoNotSaveChanges = 0
$wdCompareTargetNew = 2

# Auxiliary function
function resolve($relativePath) {
    (Resolve-Path $relativePath).Path
}

# Script body
try {
    $BaseFileName = resolve $BaseFileName
    $ChangedFileName = resolve $ChangedFileName

    # Remove the readonly attribute because Word is unable to compare readonly
    # files:
    $baseFile = Get-ChildItem $BaseFileName
    if ($baseFile.IsReadOnly) {
        $baseFile.IsReadOnly = $false
    }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $word.Documents.Open($BaseFileName, $false, $false)
    $document.Compare($ChangedFileName, [ref]"Comparison", [ref]$wdCompareTargetNew, [ref]$true, [ref]$true)

    $word.ActiveDocument.Saved = 1

    # Now close the document so only compare results window persists:
    $document.Close([ref]$wdDoNotSaveChanges)
    $word.Visible = $true
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}
