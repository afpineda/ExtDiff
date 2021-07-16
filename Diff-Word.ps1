<####################################################################

    .SYNOPSIS
        Run Ms-Word as a diff tool.
        Useful for version control.

    .PARAMETER $localFileName
        Path to first file

    .PARAMETER $remoteFileName
        Path to second file

    .NOTES
        MS-Word must be installed.

####################################################################>

param(
    [string] $localFileName,
    [string] $remoteFileName
)

$ErrorActionPreference = 'Stop'

# Constants
$wdDoNotSaveChanges = 0
$wdCompareTargetNew = 2
$wdOpenFormatAuto = 0
$wdLeftToRight = 0

<####################################################################

    .SYNOPSIS
        Auxiliary function in order to open MS-Word documents

    .PARAMETER $word
        Instance of Word.Application
    
    .PARAMETER $relativePath
        Path to file. If $null, an empty
    
    .OUTPUTS
        Document object

####################################################################>
function open-word-doc($word, $relativePath) {
    if ($relativePath) {
        $absPath = (Resolve-Path $relativePath).Path
        $file = Get-ChildItem $absPath
        $file.IsReadOnly = $false
        #$word.Documents.OpenNoRepairDialog($absPath, $false, $true)
        $word.Documents.OpenNoRepairDialog($absPath, $false, $true, $false,'','',$true,'','',$wdOpenFormatAuto,((Get-WinSystemLocale).LCID), $false,$true,$wdLeftToRight,$true)
    } else {
        $word.Documents.Add()
    }
}

<####################################################################
    Script Body
####################################################################>

try {
    # if MS Word is not installed, fail as soon as possible
    $word = New-Object -ComObject Word.Application

    # Keep MS-word invisible to prevent user from tampering while running this script
    $word.Visible = $false

    # Open documents
    $docLocal = open-word-doc $word $localFileName
    $docRemote = open-word-doc $word $remoteFileName

    # Open comparison window
    $docComparison = $word.Application.CompareDocuments($docLocal,$docRemote,$wdCompareTargetNew,1)
    $docComparison.Saved  = $true
    
    #close single documents
    $docLocal.Close([ref]$wdDoNotSaveChanges)
    $docRemote.Close([ref]$wdDoNotSaveChanges)

    # show
    $word.Visible = $true
    exit 0
} catch {
    if ($word) {
        # avoid hidden MS-word instances after failing
        $word.Quit([ref]$false)
    }
    $line = $_.InvocationInfo.ScriptLineNumber

    # show exception text in a message box
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception,"FAILURE at line $line")
    exit 1
}
