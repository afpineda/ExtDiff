#.\Merge-Word.ps1 "C:\Temp\PruebasGIT\Test1\test.docx" "C:\Temp\PruebasGIT\Test2\test.docx"
# $d1 = $word.Documents.Open("D:\Documentos\dev\ExtDiff\testCases\d1.docx")
# $d2 = $word.Documents.Open("D:\Documentos\dev\ExtDiff\testCases\d2.docx")
param(
    [string] $localFileName,
    [string] $remoteFileName,
    [string] $destinationFileName
)

$ErrorActionPreference = 'Stop'

# Constants
$wdDoNotSaveChanges = 0
$wdCompareTargetNew = 2

# Auxiliary function
function open-word-doc($word, $relativePath) {
    if ($relativePath) {
        $absPath = (Resolve-Path $relativePath).Path
        $file = Get-ChildItem $absPath
        $file.IsReadOnly = $false
        $word.Documents.Open($absPath, $false, $false)
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
    if ($destinationFileName) {
        $docComparison.SaveAs($destinationFileName)
        $startTimestamp = (Get-ChildItem $destinationFileName).LastWriteTime
    }

    #close single documents
    $docLocal.Close([ref]$wdDoNotSaveChanges)
    $docRemote.Close([ref]$wdDoNotSaveChanges)

    # show
    $word.Visible = $true

    # Wait for MS-word to close
    do {
        sleep(1)
    } while(!!$word.Application) 

    # test if destination file changed and give exit code
    if ($destinationFileName) {
        $currentTimestamp = (Get-ChildItem $destinationFileName).LastWriteTime
        if ($currentTimestamp -eq $startTimestamp) {
            exit 128
        } else {
            exit 0
        }
    }
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