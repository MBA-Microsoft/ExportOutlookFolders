param (
    [Parameter(Mandatory=$true)][string]$userEmail 
)

$OUTLOOK = new-object -comobject outlook.application
$NAMESP = $OUTLOOK.GetNameSpace("MAPI")
$SFOLDER = $NAMESP.Folders[$userEmail]
$OutputFileName = ".\OutlookFolders_" + (Get-Date).ToString('yyyyMMdd') + ".txt"

function get-allSubFolders($FOLDER_)
  {
  foreach ( $ITEM in $FOLDER_.Folders)
    {
      Write-Output $ITEM.FolderPath 
      get-allSubFolders $Item
    }
  }


get-allSubFolders $SFOLDER | Out-File $OutputFileName
