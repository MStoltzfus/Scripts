Add-Type -AssemblyName System.Windows.Forms

Clear-Host

#Base commands used for the following script
#Set-CsOnlineApplicationInstance -ApplicationId "341e195c-b261-4b05-8ba5-dd4a89b1f3e7" -Identity "Resource Account Username" - this command assigns the Landis CC Bot Application ID to a Teams Resource Account

#initialize a Filebrowser through .Net in order to pick the appropriate ingest file
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop') 
    Filter           = 'CSV File (*.csv)|*.csv'
}

Read-Host -Prompt "The next step will show you a dialog box to select the CSV file of Resource Accounts you wish to assign to Landis Contact Center. Press ENTER to continue."

$FileBrowser.ShowDialog()

Clear-Host

$path = $FileBrowser.FileName
$csv = Import-csv -path $path
$landisAppID = "341e195c-b261-4b05-8ba5-dd4a89b1f3e7"
$editedResourceAccountObjectIds = @()

import-module MicrosoftTeams
connect-MicrosoftTeams

#loops through the selected CSV file, assigns each listed UserPrincipalName to the Landis Application ID, and adds the ObjectID of the Resource Account to an array.
foreach ($line in $csv) {
    $raUserPrincipalName = $line.UserPrincipalName
    $raObjectId = (Get-CsOnlineApplicationInstance -identity $raUserPrincipalName).objectId
    Remove-CsOnlineApplicationInstanceAssociation -Identities $raObjectId | Out-Null
    Set-CsOnlineApplicationInstance -Identity $line.UserPrincipalName -ApplicationId $landisAppID | Out-Null

    $editedResourceAccountObjectIds += $raObjectId
}

Read-Host -Prompt "If you have the CSV file with the Resource Accounts List open, please close it. Press ENTER to continue."


Get-CsOnlineApplicationInstance | Select-Object displayname, userprincipalname, phoneNumber, objectId, applicationId | Where-Object { $_.ApplicationId -eq "341e195c-b261-4b05-8ba5-dd4a89b1f3e7" } | Export-Csv $path -NoTypeInformation

Disconnect-MicrosoftTeams

Clear-Host

Write-Host "Resource Accounts from the CSV file have been assigned to Landis Contact Center"

$editedResourceAccountObjectIds

& $path

Get-PSSession | Remove-PSSession