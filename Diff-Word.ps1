param(
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory=$true)]
    [string] $BaseFileName,
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
    # if MS Word is not installed, fail as soon as possible
    $word = New-Object -ComObject Word.Application

    # Keep MS-word invisible to prevent user from tampering while running this script
    $word.Visible = $false

    $BaseFileName = resolve $BaseFileName
    # Remove the readonly attribute because Word is unable to compare readonly
    # files:
    $baseFile = Get-ChildItem $BaseFileName
    if ($baseFile.IsReadOnly) {
        $baseFile.IsReadOnly = $false
    }
    # Open first document
    $docBase = $word.Documents.Open($BaseFileName, $false, $false)

    # Open second document
    # NOTE: if not given, use an empty document
    if ($ChangedFileName) {
        $ChangedFileName = resolve $ChangedFileName
        $docChanged = $word.Documents.Open($ChangedFileName, $false, $false)
    } else {
        $docChanged = $word.Documents.Add()
    }

    # Open comparison window
    $docComparison = $word.Application.CompareDocuments($docChanged,$docBase,$wdCompareTargetNew,1)
    $docComparison.Saved  = $true
    
    #close single documents
    $docBase.Close([ref]$wdDoNotSaveChanges)
    $docChanged.Close([ref]$wdDoNotSaveChanges)

    # show
    $word.Visible = $true
    exit 0
} catch {
    if ($word) {
        # avoid hidden MS-word instances after failing
        $word.Quit([ref]$false)
    }
    # show exception text in a message box
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception,"FAILURE")
    exit 1
}
