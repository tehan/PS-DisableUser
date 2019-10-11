#Paramteres to collect user data
[CmdletBinding()]
param(
    [Parameter(Mandatory=$True,Position=1)]
    [String]$Name,
    [Parameter(Mandatory=$True,Position=2)]
    [String]$Surname
)

#Variable that create user
$user = $name + '.' + $surname

#Getting current date
$Date = Get-Date -format g

#Path for logs and credentials
$DirPath = "C:\PowershellScripts\DisableUser"

#Check if DirPatch exist
$DirPathTest = Test-Path -Path $DirPath

If (!($DirPathTest)) {
    Try{
        New-Item -ItemType Directory $DirPath 
    }
    Catch {
        $_ | Out-File ($DirPatch + "\" + "log.txt") -Append
    }
}

#Credentials container for AD
$CredentialsContainerAD = ($DirPath + "\" + "DisableUserAD.cred")

#Checking if file already exist
$CheckCredenialAD = Test-Path -Path $CredentialsContainerAD

If (!($CheckCredenialAD)) {
    #Adding information to Log file when credentials has been provided.
    "$Date - Creating the Credentials Container for AD" | Out-File ($DirPath + "\" + "Log.txt") -Append
    #Gathering credentials
    $CredentialsAD = get-credential -Message "Please provide credentials that allow to disable users - Global Domain Administrator"
    #Creating encrypted file with credentials
    $CredentialsAD | Export-CliXml -Path $CredentialsContainerAD
}

#Imporitng credentials
Try {
write-host "Importing Credentials for AD..." -ForegroundColor Yellow
$credAD = (Import-CliXml -Path $CredentialsContainerAD)
write-host "Credentials for AD has been imported sucessfully" -ForegroundColor Green
}
Catch {
$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
write-host "Importing Credentials for AD has failed. Check the log file." -ForegroundColor Red
Write-host "The End" -ForegroundColor DarkCyan
Return
}

#Finding the user
Try {
write-host "Seeking for user $user..." -ForegroundColor Yellow
$UserSAM = Get-ADUser -Identity $user | Select-Object -ExpandProperty SamAccountName
write-host "Found $user..." -ForegroundColor Green
}
Catch {
$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
write-host "Cannot find $User" -ForegroundColor Red
return
}


#Loop to confirm action
$Confirmation = Read-Host "Are you sure that you want disable $user ? [y/n]"
while($Confirmation -ne "y")
{
    if ($Confirmation -eq 'n') {
    write-host "$user has not been disable" -ForegroundColor Yellow
    return
    }
    $Confirmation = Read-Host "Are you sure that you want disable $user ? [y/n]"
}

#Changing user field, disabling account, removing AD groups
Try {
    write-host "Disabling user..." -ForegroundColor Yellow
    Disable-ADAccount -Identity $UserSAM -Credential $credAD
    write-host "Removing Office and Company fields..." -ForegroundColor Yellow
    Set-AdUser $UserSam -Clear Company -Credential $credAD
    Set-AdUser $UserSam -Office $null -Credential $credAD
    write-host "Changing DisplayName..." -ForegroundColor Yellow
    $ADname = Get-ADUser -Identity $UserSAM -Properties * | Select-Object -ExpandProperty GivenName
    $Adsurname = Get-ADUser -Identity $UserSAM -Properties * | Select-Object -ExpandProperty Surname
    $DisplayName = 'z_' + $ADname + ' ' + $ADsurname
    Set-AdUser $UserSam -DisplayName $DisplayName -Credential $credAD
    write-host "Set Description as current date..." -ForegroundColor Yellow
    Set-AdUser $UserSam -Description $Date -Credential $credAD
    write-host "Removing Groups Membership..." -ForegroundColor Yellow
    Get-AdUser -Identity $UserSam -Properties memberof | Select-Object -ExpandProperty memberof | Remove-ADGroupMember -Members $UserSAM -Credential $credAD -confirm:$false
}
Catch {
    $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
    write-host "Something went wrong, check the log file" -ForegroundColor Red
    return
}
write-host "You just disable $user" -ForegroundColor DarkCyan

#Forcing synchornization between Domain Controllers 
$DCscriptblock = { repadmin /syncall nvmdc03 "dc=general,dc=newvoicemedia,dc=com" /force } 
try {
    write-host "Synchronization between Domain Controllers..." -ForegroundColor Yellow
    Invoke-Command -computername nvmdc03 -Credential $credAD -scriptblock $DCscriptblock | Out-File ($DirPath + "\" + "Log.txt") -Append
    write-host "Synchronization between Domain Controllers completed successfully." -ForegroundColor Green
}
catch {
     $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
     write-host "Synchronization between Domain Controllers failed." -ForegroundColor Red
}


#Loop to confirm if you would like procced to Office 365
$Office365 = Read-Host "Would you like to block user in Office 365? [y/n]"
while($Office365 -ne "y")
{
    if ($Office365 -eq 'n') {
    Write-host "The End" -ForegroundColor DarkCyan
    return
    }
    $Office365 = Read-Host "Would you like to block user in Office 365? [y/n]"
}

#Credentials container for Office365
$CredentialsContainerO365 = ($DirPath + "\" + "DisableUserO365.cred")

#Checking if file already exist
$CheckCredenialO365 = Test-Path -Path $CredentialsContainerO365

If (!($CheckCredenialO365)) {
    #Adding information to Log file when credentials has been provided.
    "$Date - Creating the Credentials Container for O365" | Out-File ($DirPath + "\" + "Log.txt") -Append
    #Gathering credentials
    $CredentialsO365 = get-credential -Message "Please provide credentials that allow to disable users - Office 365 Administrator"
    #Creating encrypted file with credentials
    $CredentialsO365 | Export-CliXml -Path $CredentialsContainerO365
}

#Imporitng credentials
Try {
write-host "Importing Credentials for O365..." -ForegroundColor Yellow
$credO365 = (Import-CliXml -Path $CredentialsContainerO365)
write-host "Credentials for O365 has been imported sucessfully" -ForegroundColor Green
}
Catch {
$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
write-host "Importing Credentials for O365 has failed. Check the log file." -ForegroundColor Red
Write-host "The End" -ForegroundColor DarkCyan
Return
}

#Connecting to Office 365 online serivce
Try {
    write-host "Connecting to MsolService..." -ForegroundColor Yellow
    Connect-MsolService -Credential $credO365
    write-host "Connected to MsolService successfully." -ForegroundColor Green
}
Catch {
    $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
    write-host "Connecting to MsolService has failed" -ForegroundColor Red
    Write-host "The End" -ForegroundColor DarkCyan
    Return
}

#Connecting to SharePoint online service
Try {
    write-host "Connecting to SPOService..." -ForegroundColor Yellow
    Connect-sposervice -url https://newvoicemedia-admin.sharepoint.com -Credential $credO365
    write-host "Connected to SPOService successfully." -ForegroundColor Green
}
Catch {
    $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
    write-host "Connecting to SPOService has failed" -ForegroundColor Red
    Write-host "The End" -ForegroundColor DarkCyan
    Return
}

#Blocking user
Try {
    Write-host "Blocking user from sign-on..." -ForegroundColor Yellow
    $UPN = Get-ADUser -Identity $user -Properties * | Select-Object -ExpandProperty UserPrincipalName
    Set-MsolUser -UserPrincipalName $UPN -BlockCredential $true
    Write-host "Changing user password..." -ForegroundColor Yellow
    Set-MsolUserPassword -UserPrincipalName $UPN -ForceChangePassword $true
    Write-Host "Removing all user sessions...." -ForegroundColor Yellow
    Revoke-SPOUserSession -user $UPN -Confirm:$false
}
Catch {
    $_ | Out-File ($DirPath + "\" + "Log.txt") -Append
    write-host "Could not block the user" -ForegroundColor Red
    Write-host "The End" -ForegroundColor DarkCyan
    Return
}


Write-host "The End" -ForegroundColor DarkCyan

