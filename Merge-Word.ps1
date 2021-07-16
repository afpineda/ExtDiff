<####################################################################

    .SYNOPSIS
        Run Ms-Word as a merge tool.
        Useful for version control.

    .PARAMETER $localFileName
        Path to original file

    .PARAMETER $remoteFileName
        Path to changed file

    .PARAMETER $mergedFileName
        Path to merged file

    .PARAMTER $Force
        If $true, file is merged with no user interaction.
        Since change tracking is enabled in merge mode, changes
        may be discarded or accepted later
  
    .NOTES
        MS-Word must be installed.
        $LASTEXITCODE is 0 if-and-only-if merged file is NOT SAVED by
        the user.

    .EXAMPLE
        .\Merge-Word.ps1 "D:\Documentos\dev\ExtDiff\testCases\d1.docx" "D:\Documentos\dev\ExtDiff\testCases\d2.docx" "E:\temp\PruebasGIT\tmp.docx"

####################################################################>


param(
    [string] $localFileName,
    [string] $remoteFileName,
    [string] $mergedFileName,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

# Constants
$wdDoNotSaveChanges = 0
$wdCompareTargetNew = 2

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
        $locale = (Get-WinSystemLocale).LCID
        $word.Documents.OpenNoRepairDialog($absPath, $false, $true, $false,'','',$true,'','',$wdOpenFormatAuto, $locale, $false,$true,$wdLeftToRight,$true)
    } else {
        $word.Documents.Add()
    }
}

# Script body
try {
    # if MS Word is not installed, fail as soon as possible
    $word = New-Object -ComObject Word.Application

    # Keep MS-word invisible to prevent user from tampering while running this script
    #$word.Visible = $true
    $word.Visible = $false

    # Open documents
    $docLocal = open-word-doc $word $localFileName
    $docRemote = open-word-doc $word $remoteFileName

    # Open merge window
    $docComparison = $word.Application.MergeDocuments($docLocal,$docRemote,$wdCompareTargetNew)
    if ($mergedFileName) {
        $docComparison.SaveAs($mergedFileName)
        $startTimestamp = (Get-ChildItem $mergedFileName).LastWriteTime
    }

    #close single documents
    $docLocal.Close([ref]$wdDoNotSaveChanges)
    $docRemote.Close([ref]$wdDoNotSaveChanges)

    if ($Force -eq $true)
    {
        # NON-INTERACTIVE MODE
        $word.Quit([ref]$false)
        exit 0

    } else {

        # INTERACTIVE MODE
        $word.Visible = $true
    
        # Wait for MS-word to close
        do {
            sleep(1)
        } while(!!$word.Application) 
    
        # test if destination file changed and give exit code
        if ($mergedFileName) {
            $currentTimestamp = (Get-ChildItem $mergedFileName).LastWriteTime
            if ($currentTimestamp -eq $startTimestamp) {
                exit 128
            } else {
                exit 0
            }
        }
    }
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